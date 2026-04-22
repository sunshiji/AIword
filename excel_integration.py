import os
import tempfile
import pythoncom
import win32com.client
from typing import Optional
from logger import Logger


class ExcelIntegration:
    def __init__(self, logger: Logger):
        self.logger = logger
    
    def _get_running_excel_app(self):
        try:
            pythoncom.CoInitialize()
            app = win32com.client.GetActiveObject("Excel.Application")
            self.logger.info("已连接到已运行的 Excel 实例")
            return app
        except Exception as e:
            self.logger.debug(f"没有找到已运行的 Excel 实例: {str(e)}")
            return None
    
    def _get_running_et_app(self):
        try:
            pythoncom.CoInitialize()
            try:
                app = win32com.client.GetActiveObject("KET.Application")
                self.logger.info("已连接到已运行的 WPS 表格实例")
                return app
            except:
                app = win32com.client.GetActiveObject("Et.Application")
                self.logger.info("已连接到已运行的 WPS 表格实例")
                return app
        except Exception as e:
            self.logger.debug(f"没有找到已运行的 WPS 表格实例: {str(e)}")
            return None
    
    def _start_excel_app(self):
        try:
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = True
            self.logger.info("已启动新的 Excel 实例")
            return app
        except Exception as e:
            self.logger.error(f"无法启动 Excel: {str(e)}")
            return None
    
    def _start_et_app(self):
        try:
            pythoncom.CoInitialize()
            try:
                app = win32com.client.Dispatch("KET.Application")
                app.Visible = True
                self.logger.info("已启动新的 WPS 表格实例")
                return app
            except:
                app = win32com.client.Dispatch("Et.Application")
                app.Visible = True
                self.logger.info("已启动新的 WPS 表格实例")
                return app
        except Exception as e:
            self.logger.error(f"无法启动 WPS 表格: {str(e)}")
            return None
    
    def get_running_app(self):
        excel_app = self._get_running_excel_app()
        if excel_app:
            return excel_app
        et_app = self._get_running_et_app()
        if et_app:
            return et_app
        return None
    
    def start_app(self, prefer_et: bool = True):
        if prefer_et:
            et_app = self._start_et_app()
            if et_app:
                return et_app
            excel_app = self._start_excel_app()
            return excel_app
        else:
            excel_app = self._start_excel_app()
            if excel_app:
                return excel_app
            et_app = self._start_et_app()
            return et_app
    
    def get_active_workbook(self, allow_start: bool = False):
        app = self.get_running_app()
        
        if not app and allow_start:
            app = self.start_app()
        
        if app:
            try:
                return app.ActiveWorkbook
            except Exception as e:
                self.logger.debug(f"获取活动工作簿失败: {str(e)}")
                pass
        
        return None
    
    def is_available(self) -> bool:
        return self.get_running_app() is not None
    
    def insert_csv(self, csv_content: str, allow_start: bool = False) -> bool:
        try:
            workbook = self.get_active_workbook(allow_start=allow_start)
            if not workbook:
                if allow_start:
                    self.logger.warning("无法启动 Excel 或 WPS 表格")
                return False
            
            with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8-sig') as f:
                f.write(csv_content)
                temp_path = f.name
            
            try:
                worksheet = workbook.ActiveSheet
                excel_app = workbook.Application
                
                existing_data = worksheet.UsedRange
                if existing_data:
                    start_row = existing_data.Row + existing_data.Rows.Count
                else:
                    start_row = 1
                
                start_cell = worksheet.Cells(start_row, 1)
                
                temp_path_abs = os.path.abspath(temp_path)
                query_table = worksheet.QueryTables.Add(
                    Connection=f"TEXT;{temp_path_abs}",
                    Destination=start_cell
                )
                query_table.TextFileParseType = 1
                query_table.TextFileCommaDelimiter = True
                query_table.TextFileColumnDataTypes = [2] * 100
                query_table.Refresh(BackgroundQuery=False)
                query_table.Delete()
                
                self.logger.info("CSV 内容已成功插入到表格")
                return True
            finally:
                if os.path.exists(temp_path):
                    try:
                        os.unlink(temp_path)
                    except:
                        pass
                    
        except Exception as e:
            self.logger.error(f"插入 CSV 内容失败: {str(e)}")
            return False
    
    def save_as_xlsx(self, csv_content: str, save_dir: str) -> Optional[str]:
        try:
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
            
            excel_app = self.get_running_app()
            if not excel_app:
                excel_app = self.start_app()
            
            if not excel_app:
                self.logger.error("无法启动 Excel 或 WPS 表格")
                return None
            
            workbook = excel_app.Workbooks.Add()
            
            with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8-sig') as f:
                f.write(csv_content)
                temp_path = f.name
            
            try:
                worksheet = workbook.ActiveSheet
                
                query_table = worksheet.QueryTables.Add(
                    Connection=f"TEXT;{os.path.abspath(temp_path)}",
                    Destination=worksheet.Cells(1, 1)
                )
                query_table.TextFileParseType = 1
                query_table.TextFileCommaDelimiter = True
                query_table.TextFileColumnDataTypes = [2] * 100
                query_table.Refresh(BackgroundQuery=False)
                query_table.Delete()
                
                timestamp = self._get_timestamp()
                filename = f"AIword_table_{timestamp}.xlsx"
                filepath = os.path.join(save_dir, filename)
                
                workbook.SaveAs(os.path.abspath(filepath))
                workbook.Close()
                
                self.logger.info(f"表格已保存到: {filepath}")
                return filepath
            finally:
                if os.path.exists(temp_path):
                    try:
                        os.unlink(temp_path)
                    except:
                        pass
                    
        except Exception as e:
            self.logger.error(f"保存表格失败: {str(e)}")
            return None
    
    def _get_timestamp(self) -> str:
        from datetime import datetime
        return datetime.now().strftime('%Y%m%d_%H%M%S')
