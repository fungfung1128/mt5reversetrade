import sys
import socket
import os
import logging
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                           QHBoxLayout, QLabel, QPushButton, QTableWidget,
                           QTableWidgetItem, QHeaderView, QTabWidget)
from PyQt6.QtCore import Qt, QTimer
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import MetaTrader5 as mt5
from datetime import datetime
import numpy

# 設置日誌檔案
logging.basicConfig(filename='rtrade.log', level=logging.INFO, encoding='utf-8')

class MT5TradeGenerator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("XAUUSD交易指令生成器")
        self.setGeometry(100, 100, 900, 500)

        # 定義產品名稱映射
        self.mt5_symbol = "XAUUSD.ECN"
        self.google_symbol = "xauusd"
        self.internal_symbol = "XAUUSD"

        # 初始化 Google Sheets 客戶端
        self.gc = None
        self.worksheet = None
        self.spreadsheet = None

        # 初始化 MT5 連線狀態
        self.mt5_connected = False

        # 主界面組件
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        # 標籤頁
        self.tab_widget = QTabWidget()
        self.layout.addWidget(self.tab_widget)

        # 交易頁
        self.trade_widget = QWidget()
        self.trade_layout = QVBoxLayout()
        self.trade_widget.setLayout(self.trade_layout)
        self.tab_widget.addTab(self.trade_widget, "交易")

        # 標題
        self.title_label = QLabel("XAUUSD交易指令生成器")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        self.trade_layout.addWidget(self.title_label)

        # 狀態標籤
        self.status_label = QLabel("狀態: 未連接到 MT5 和 Google Sheets")
        self.trade_layout.addWidget(self.status_label)

        # 連線按鈕
        self.connect_button = QPushButton("連接到 MT5 和 Google Sheets")
        self.connect_button.clicked.connect(self.connect_to_mt5_and_google_sheets)
        self.trade_layout.addWidget(self.connect_button)

        # 數據表格
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["產品", "當前手數(MT5)", "Google淨手數", "目標手數", "交易指令"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.trade_layout.addWidget(self.table)

        # 操作按鈕
        self.button_layout = QHBoxLayout()
        self.refresh_button = QPushButton("刷新數據")
        self.refresh_button.clicked.connect(self.refresh_data)
        self.refresh_button.setEnabled(False)
        self.button_layout.addWidget(self.refresh_button)

        self.generate_button = QPushButton("生成交易指令")
        self.generate_button.clicked.connect(self.generate_trades)
        self.generate_button.setEnabled(False)
        self.button_layout.addWidget(self.generate_button)

        self.execute_button = QPushButton("執行交易 (真實)")
        self.execute_button.clicked.connect(self.execute_trades)
        self.execute_button.setEnabled(False)
        self.button_layout.addWidget(self.execute_button)
        self.trade_layout.addLayout(self.button_layout)

        # 自動刷新定時器
        self.refresh_timer = QTimer()
        self.refresh_timer.timeout.connect(self.refresh_data)
        self.auto_refresh = False

        # 自動刷新復選框
        self.auto_refresh_checkbox = QPushButton("啟用自動刷新 (10秒)")
        self.auto_refresh_checkbox.setCheckable(True)
        self.auto_refresh_checkbox.clicked.connect(self.toggle_auto_refresh)
        self.auto_refresh_checkbox.setEnabled(False)
        self.trade_layout.addWidget(self.auto_refresh_checkbox)

        # 自動交易復選框
        self.auto_trade_checkbox = QPushButton("啟用自動交易")
        self.auto_trade_checkbox.setCheckable(True)
        self.auto_trade_checkbox.clicked.connect(self.toggle_auto_trade)
        self.auto_trade_checkbox.setEnabled(False)
        self.trade_layout.addWidget(self.auto_trade_checkbox)
        self.auto_trade = False

        # 日誌頁
        self.log_widget = QWidget()
        self.log_layout = QVBoxLayout()
        self.log_widget.setLayout(self.log_layout)
        self.log_table = QTableWidget()
        self.log_table.setColumnCount(2)
        self.log_table.setHorizontalHeaderLabels(["時間", "日誌信息"])
        self.log_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.log_layout.addWidget(self.log_table)
        self.tab_widget.addTab(self.log_widget, "日誌")

        # 初始化數據
        self.current_positions = {}
        self.google_positions = {}
        self.last_trade_time = None
        self.zero_check_count = 0
        self.zero_check_timer = QTimer()
        self.zero_check_timer.timeout.connect(self.verify_zero_position)
        self.last_non_zero_lot = None

        # 自動連接到 MT5
        self.connect_to_mt5_and_fetch_positions()

    def connect_to_mt5_and_fetch_positions(self):
        try:
            if not mt5.initialize():
                error_msg = f"MT5 初始化失敗，錯誤代碼: {mt5.last_error()}"
                self.log_message(f"錯誤: {error_msg}")
                print(error_msg)
                self.status_label.setText("狀態: MT5 連線失敗")
                return

            self.mt5_connected = True
            account_info = mt5.account_info()
            print(f"MT5 連線成功，帳戶: {account_info.login}")
            self.log_message(f"信息: MT5 連線成功，帳戶: {account_info.login}")

            self.update_mt5_positions()
            self.update_table()
            self.status_label.setText("狀態: 已連接到 MT5，等待 Google Sheets 連線")
            self.refresh_button.setEnabled(True)
            self.generate_button.setEnabled(True)
            self.execute_button.setEnabled(True)
            self.auto_refresh_checkbox.setEnabled(True)
            self.auto_trade_checkbox.setEnabled(True)

        except Exception as e:
            error_msg = f"MT5 連線錯誤: {str(e)}"
            self.log_message(f"錯誤: {error_msg}")
            print(error_msg)
            self.status_label.setText("狀態: MT5 連線失敗")
            self.mt5_connected = False

    def connect_to_mt5_and_google_sheets(self):
        try:
            socket.setdefaulttimeout(30)
            scope = [
                'https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive'
            ]
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
            json_path = os.path.join(base_path, 'impactful-name-455509-b6-b07e866843f7.json')
            creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)
            self.gc = gspread.authorize(creds)
            print(f"服務帳號: {creds.service_account_email}")
            self.log_message(f"信息: 服務帳號: {creds.service_account_email}")

            try:
                self.spreadsheet = self.gc.open("data")
                print("可用工作表:", [sheet.title for sheet in self.spreadsheet.worksheets()])
                self.worksheet = self.spreadsheet.worksheet("Net Position")
                print(f"已連線工作表: {self.worksheet.title}")
                self.log_message(f"信息: 已連線工作表: {self.worksheet.title}")
            except Exception as e:
                error_msg = f"無法訪問工作表: {str(e)}"
                self.log_message(f"錯誤: {error_msg}")
                return

            self.status_label.setText("狀態: 已連接到 MT5 和 Google Sheets")
            self.connect_button.setEnabled(False)
            self.refresh_data()

        except Exception as e:
            error_msg = f"Google Sheets 連線錯誤: {str(e)}"
            self.log_message(f"錯誤: {error_msg}")
            print(error_msg)
            self.status_label.setText("狀態: Google Sheets 連線失敗")

    def update_mt5_positions(self):
        positions = mt5.positions_get(symbol=self.mt5_symbol)
        self.current_positions = {}
        if positions:
            net_lots = 0.0
            for pos in positions:
                if pos.symbol == self.mt5_symbol:
                    lots = pos.volume if pos.type == mt5.ORDER_TYPE_BUY else -pos.volume
                    net_lots += lots
                    print(f"MT5 持倉: {pos.symbol}, 類型: {'買入' if pos.type == mt5.ORDER_TYPE_BUY else '賣出'}, 手數: {pos.volume}")
            self.current_positions[self.internal_symbol] = net_lots
            print(f"MT5 淨持倉: {self.internal_symbol}, 手數: {net_lots}")
            self.log_message(f"信息: MT5 淨持倉: {self.internal_symbol}, 手數: {net_lots}")
        else:
            self.current_positions[self.internal_symbol] = 0.0
            print(f"MT5 淨持倉: {self.internal_symbol}, 手數: 0.0")
            self.log_message(f"信息: MT5 淨持倉: {self.internal_symbol}, 手數: 0.0")

    def close_opposite_positions(self, symbol, desired_action, desired_lots):
        positions = mt5.positions_get(symbol=symbol)
        if not positions:
            return 0.0

        total_closed_lots = 0.0
        for pos in positions:
            if pos.symbol != symbol:
                continue

            if (desired_action == mt5.ORDER_TYPE_BUY and pos.type == mt5.ORDER_TYPE_SELL) or \
               (desired_action == mt5.ORDER_TYPE_SELL and pos.type == mt5.ORDER_TYPE_BUY):
                close_request = {
                    "action": mt5.TRADE_ACTION_DEAL,
                    "position": pos.ticket,
                    "symbol": symbol,
                    "volume": pos.volume,
                    "type": mt5.ORDER_TYPE_BUY if pos.type == mt5.ORDER_TYPE_SELL else mt5.ORDER_TYPE_SELL,
                    "price": mt5.symbol_info_tick(symbol).ask if pos.type == mt5.ORDER_TYPE_SELL else mt5.symbol_info_tick(symbol).bid,
                    "type_time": mt5.ORDER_TIME_GTC,
                    "type_filling": mt5.ORDER_FILLING_IOC,
                }

                while True:
                    result = mt5.order_send(close_request)
                    if result.retcode == mt5.TRADE_RETCODE_DONE:
                        total_closed_lots += pos.volume
                        self.log_message(f"信息: 已平倉相反持倉 - {symbol}, 手數: {pos.volume}")
                        break
                    elif result.retcode == mt5.TRADE_RETCODE_REQUOTE:
                        self.log_message(f"警告: 平倉時出現 Requote，重新以新價格 {result.price} 執行")
                        close_request["price"] = result.price
                        continue
                    else:
                        self.log_message(f"錯誤: 平倉失敗，錯誤代碼: {result.retcode}, 詳情: {result.comment}")
                        break

        return total_closed_lots

    def log_message(self, message):
        logging.info(message)
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        row_count = self.log_table.rowCount()
        self.log_table.insertRow(row_count)
        self.log_table.setItem(row_count, 0, QTableWidgetItem(current_time))
        self.log_table.setItem(row_count, 1, QTableWidgetItem(message))
        self.log_table.scrollToBottom()

    def verify_zero_position(self):
        self.zero_check_count += 1
        try:
            all_data = self.worksheet.get_all_values()
            for row_idx, row in enumerate(all_data):
                if len(row) >= 3 and str(row[1]).strip().lower() == self.google_symbol:
                    lot_str = str(row[2]).strip()
                    if lot_str == "":
                        lot = 0.0
                    else:
                        try:
                            clean_lot = lot_str.replace(',', '').replace(' ', '')
                            lot = float(clean_lot)
                        except ValueError:
                            self.log_message(f"錯誤: 行 {row_idx + 1} {self.google_symbol} 手數格式無效: '{lot_str}'")
                            return

                    if lot != 0.0:
                        self.google_positions[self.internal_symbol] = lot
                        self.last_non_zero_lot = lot
                        self.zero_check_timer.stop()
                        self.zero_check_count = 0
                        self.log_message(f"信息: 檢測到非 0 值 ({lot})，停止 0 值檢查")
                        self.update_table()
                        if self.auto_trade:
                            self.execute_trades()
                        return

            if self.zero_check_count >= 3:
                self.google_positions[self.internal_symbol] = 0.0
                self.zero_check_timer.stop()
                self.zero_check_count = 0
                self.log_message(f"信息: 連續三次檢測到 0，確認 Google Sheets 持倉為 0")
                self.update_table()
                if self.auto_trade:
                    self.execute_trades()
            else:
                self.log_message(f"信息: 第 {self.zero_check_count} 次檢測到 0，等待下一次檢查")
        except Exception as e:
            self.log_message(f"錯誤: 驗證 0 值時出錯: {str(e)}")
            self.zero_check_timer.stop()
            self.zero_check_count = 0

    def refresh_data(self):
        if not self.worksheet or not self.mt5_connected:
            self.log_message("錯誤: 未連接到 MT5 或未找到有效的工作表")
            return

        try:
            print("\n------ 開始刷新數據 ------")
            self.log_message("信息: 開始刷新數據")
            self.update_mt5_positions()

            all_data = self.worksheet.get_all_values()
            self.google_positions = {}
            xauusd_found = False

            for row_idx, row in enumerate(all_data):
                if len(row) >= 3:
                    product = str(row[1]).strip()
                    lot_str = str(row[2]).strip()
                    if product.lower() == self.google_symbol:
                        if lot_str == "":
                            if self.last_non_zero_lot is not None:
                                self.zero_check_count = 1
                                self.log_message(f"信息: 檢測到空值，啟動 0 值驗證 (第 1 次)")
                                self.zero_check_timer.start(3000)
                                return
                            else:
                                self.google_positions[self.internal_symbol] = 0.0
                                xauusd_found = True
                                print(f"找到 {self.google_symbol} 數據 - 行 {row_idx + 1}: 產品='{product}', 手數=0.0 (空格)")
                                self.log_message(f"信息: 找到 {self.google_symbol} 數據 - 行 {row_idx + 1}: 產品='{product}', 手數=0.0 (空格)")
                                break
                        try:
                            clean_lot = lot_str.replace(',', '').replace(' ', '')
                            lot = float(clean_lot)
                            self.google_positions[self.internal_symbol] = lot
                            self.last_non_zero_lot = lot
                            xauusd_found = True
                            self.zero_check_count = 0
                            self.zero_check_timer.stop()
                            print(f"找到 {self.google_symbol} 數據 - 行 {row_idx + 1}: 產品='{product}', 手數={lot}")
                            self.log_message(f"信息: 找到 {self.google_symbol} 數據 - 行 {row_idx + 1}: 產品='{product}', 手數={lot}")
                            break
                        except ValueError:
                            print(f"行 {row_idx + 1} {self.google_symbol} 手數格式無效: '{lot_str}'")
                            self.log_message(f"錯誤: 行 {row_idx + 1} {self.google_symbol} 手數格式無效: '{lot_str}'")

            if not xauusd_found:
                print("工作表中所有產品名稱:")
                for row_idx, row in enumerate(all_data):
                    if len(row) >= 2:
                        print(f"行 {row_idx + 1}: '{row[1]}'")
                if self.last_non_zero_lot is not None:
                    self.zero_check_count = 1
                    self.log_message(f"信息: 未找到 {self.google_symbol}，啟動 0 值驗證 (第 1 次)")
                    self.zero_check_timer.start(3000)
                    return
                else:
                    self.google_positions[self.internal_symbol] = 0.0
                    self.log_message(f"警告: 在工作表中未找到小寫 '{self.google_symbol}' 產品或手數為空格，假設 Google Sheets 持倉為 0")

            self.update_table()
            if self.auto_trade:
                self.execute_trades()

            self.status_label.setText(f"狀態: 已加載 {self.google_symbol} 數據" if xauusd_found else f"狀態: 未找到 {self.google_symbol}，假設持倉為 0")
            self.log_message(f"信息: 狀態: {'已加載 ' + self.google_symbol + ' 數據' if xauusd_found else '未找到 ' + self.google_symbol + '，假設持倉為 0'}")

        except Exception as e:
            error_msg = f"刷新數據時出錯: {str(e)}"
            self.log_message(f"錯誤: {error_msg}")
            print(error_msg)
            self.status_label.setText("狀態: 刷新失敗")

    def update_table(self):
        print("正在更新表格...")
        self.log_message("信息: 正在更新表格")
        print(f"當前持倉(MT5): {self.current_positions}")
        print(f"Google持倉: {self.google_positions}")

        self.table.setRowCount(len(self.google_positions) or 1)
        row = 0
        if not self.google_positions:
            product_item = QTableWidgetItem(self.internal_symbol)
            product_item.setFlags(product_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 0, product_item)

            current_item = QTableWidgetItem(f"{self.current_positions.get(self.internal_symbol, 0.0):.2f}")
            current_item.setFlags(current_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 1, current_item)

            google_item = QTableWidgetItem("0.00")
            google_item.setFlags(google_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 2, google_item)

            target_item = QTableWidgetItem("0.00")
            target_item.setFlags(target_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 3, target_item)

            instruction_item = QTableWidgetItem("無操作")
            instruction_item.setFlags(instruction_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 4, instruction_item)
        else:
            for product, google_lot in self.google_positions.items():
                current_lot = self.current_positions.get(product, 0.0)
                target_lot = -google_lot

                product_item = QTableWidgetItem(product)
                product_item.setFlags(product_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, 0, product_item)

                current_item = QTableWidgetItem(f"{current_lot:.2f}")
                current_item.setFlags(current_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, 1, current_item)

                google_item = QTableWidgetItem(f"{google_lot:.2f}")
                google_item.setFlags(google_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, 2, google_item)

                target_item = QTableWidgetItem(f"{target_lot:.2f}")
                target_item.setFlags(target_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, 3, target_item)

                trade_instruction = self.calculate_trade_instruction(current_lot, google_lot)
                instruction_item = QTableWidgetItem(trade_instruction)
                instruction_item.setFlags(instruction_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, 4, instruction_item)

                row += 1

        print("表格更新完成")
        self.log_message("信息: 表格更新完成")

    def calculate_trade_instruction(self, current, google_lot):
        desired_mt5_position = -google_lot
        difference = desired_mt5_position - current
        if abs(difference) < 0.01:
            return "無操作"
        elif difference > 0:
            return f"買入 {abs(difference):.2f} 手"
        else:
            return f"賣出 {abs(difference):.2f} 手"

    def generate_trades(self):
        if not self.google_positions:
            self.log_message(f"警告: 未找到 {self.google_symbol} 交易數據，假設 Google Sheets 持倉為 0")
            self.google_positions[self.internal_symbol] = 0.0

        trade_instructions = []
        for product, google_lot in self.google_positions.items():
            current_lot = self.current_positions.get(product, 0.0)
            desired_mt5_position = -google_lot
            difference = desired_mt5_position - current_lot
            if abs(difference) >= 0.01:
                action = "BUY" if difference > 0 else "SELL"
                lots = abs(difference)
                trade_instructions.append(f"{product}: {action} {lots:.2f} lots at market price")

        if trade_instructions:
            msg = f"{self.google_symbol} 交易指令 (反向):\n" + "\n".join(trade_instructions)
            self.log_message(f"信息: 交易指令生成成功:\n{msg}")
        else:
            self.log_message(f"信息: {self.google_symbol} 當前無需交易")

    def execute_trades(self):
        if not self.google_positions or not self.mt5_connected:
            self.log_message(f"警告: 未連接到 MT5 或未找到 {self.google_symbol} 交易數據")
            return

        current_time = datetime.now()
        if self.last_trade_time and (current_time - self.last_trade_time).total_seconds() < 10:
            print("交易頻率過高，需等待 10 秒")
            self.log_message("警告: 交易頻率過高，需等待 10 秒")
            return

        executed_trades = []
        for product, google_lot in self.google_positions.items():
            current_lot = self.current_positions.get(product, 0.0)
            desired_mt5_position = -google_lot
            difference = desired_mt5_position - current_lot

            if abs(difference) < 0.01:
                continue

            action = mt5.ORDER_TYPE_BUY if difference > 0 else mt5.ORDER_TYPE_SELL
            lots = abs(difference)
            symbol = self.mt5_symbol

            symbol_info = mt5.symbol_info_tick(symbol)
            if not symbol_info:
                error_msg = f"無法獲取 {symbol} 的市場價格"
                self.log_message(f"錯誤: {error_msg}")
                print(error_msg)
                continue

            self.log_message(f"信息: 當前持倉: {current_lot}, 目標持倉: {desired_mt5_position}, 需要交易: {difference}")

            closed_lots = self.close_opposite_positions(symbol, action, lots)
            self.update_mt5_positions()

            current_lot = self.current_positions.get(product, 0.0)
            difference = desired_mt5_position - current_lot
            if abs(difference) < 0.01:
                continue

            lots = abs(difference)
            action = mt5.ORDER_TYPE_BUY if difference > 0 else mt5.ORDER_TYPE_SELL
            self.log_message(f"信息: 調整後持倉: {current_lot}, 最終交易: {'買入' if action == mt5.ORDER_TYPE_BUY else '賣出'} {lots:.2f} 手")

            request = {
                "action": mt5.TRADE_ACTION_DEAL,
                "symbol": symbol,
                "volume": lots,
                "type": action,
                "price": symbol_info.ask if action == mt5.ORDER_TYPE_BUY else symbol_info.bid,
                "type_time": mt5.ORDER_TIME_GTC,
                "type_filling": mt5.ORDER_FILLING_IOC,
            }

            while True:
                result = mt5.order_send(request)
                if result.retcode == mt5.TRADE_RETCODE_DONE:
                    executed_trades.append(f"{product}: {'買入' if action == mt5.ORDER_TYPE_BUY else '賣出'} {lots:.2f} 手 (真實) @ {current_time}")
                    self.log_message(f"信息: 交易成功 - {product}: {'買入' if action == mt5.ORDER_TYPE_BUY else '賣出'} {lots:.2f} 手 @ {current_time}")
                    break
                elif result.retcode == mt5.TRADE_RETCODE_REQUOTE:
                    self.log_message(f"警告: 出現 Requote，重新以新價格 {result.price} 執行")
                    request["price"] = result.price
                    continue
                else:
                    error_msg = f"交易失敗，錯誤代碼: {result.retcode}, 詳情: {result.comment}"
                    self.log_message(f"錯誤: {error_msg}")
                    print(error_msg)
                    break

            self.update_mt5_positions()

        if executed_trades:
            msg = f"已執行 {self.google_symbol} 交易 (反向, 真實):\n" + "\n".join(executed_trades)
            self.log_message(f"信息: 交易執行成功:\n{msg}")
            self.last_trade_time = current_time
            self.update_table()
        else:
            self.log_message(f"信息: 無需執行交易")

    def toggle_auto_refresh(self):
        self.auto_refresh = not self.auto_refresh
        if self.auto_refresh:
            self.refresh_timer.start(10000)
            self.auto_refresh_checkbox.setText("禁用自動刷新")
            self.auto_refresh_checkbox.setStyleSheet("background-color: lightgreen")
            self.log_message("信息: 已啟用自動刷新")
            print("已啟用自動刷新")
        else:
            self.refresh_timer.stop()
            self.auto_refresh_checkbox.setText("啟用自動刷新 (10秒)")
            self.auto_refresh_checkbox.setStyleSheet("")
            self.log_message("信息: 已禁用自動刷新")
            print("已禁用自動刷新")

    def toggle_auto_trade(self):
        self.auto_trade = not self.auto_trade
        if self.auto_trade:
            self.auto_trade_checkbox.setText("禁用自動交易")
            self.auto_trade_checkbox.setStyleSheet("background-color: lightcoral")
            self.log_message("信息: 已啟用自動交易")
            print("已啟用自動交易")
        else:
            self.auto_trade_checkbox.setText("啟用自動交易")
            self.auto_trade_checkbox.setStyleSheet("")
            self.log_message("信息: 已禁用自動交易")
            print("已禁用自動交易")

    def closeEvent(self, event):
        if self.mt5_connected:
            mt5.shutdown()
            self.log_message("信息: MT5 連線已關閉")
            print("MT5 連線已關閉")
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MT5TradeGenerator()
    window.show()
    sys.exit(app.exec())