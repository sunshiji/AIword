import os
import tempfile
import pythoncom
import win32com.client
from typing import Optional
from logger import Logger


class ExcelIntegration:
    def __init__(self, logger: Logger):
        self.logger = logger
    
    def _get_excel_app(self):
        try:
            pythoncom.CoInitialize()
            try:
                app = win32com.client.GetActiveObject("Excel.Application")
                self.logger.info("已连接到已运行的 Excel 实例")
                return app
            except:
                app = win32com.client.Dispatch("Excel.Application")
                app.Visible = True
                self.logger.info("已启动新的 Excel 实例")
                return app
        except Exception as e:
            self.logger.error(f"无法连接到 Excel: {str(e)}")
            return None
    
    def _get_et_app(self):
        try:
            pythoncom.CoInitialize()
            try:
                app = win32com.client.GetActiveObject("KET.Application")
                self.logger.info("已连接到已运行的 WPS 表格实例")
                return app
            except:
                try:
                    app = win32com.client.GetActiveObject("Et.Application")
                    self.logger.info("已连接到已运行的 WPS 表格实例")
                    return app
                except:
                    app = win32com.client.Dispatch("KET.Application")
                    app.Visible = True
                    self.logger.info("已启动新的 WPS 表格实例")
                    return app
        except Exception as e:
            self.logger.error(f"无法连接到 WPS 表格: {str(e)}")
            return None
    
    def get_active_workbook(self):
        excel_app = self._get_excel_app()
        if excel_app:
            try:
                return excel_app.ActiveWorkbook
            except:
                pass
        
        et_app = self._get_et_app()
        if et_app:
            try:
                return et_app.ActiveWorkbook
            except:
                pass
        
        return None
    
    def is_available(self) -> bool:
        return self.get_active_workbook() is not None
    
    def insert_csv(self, csv_content: str) -> bool:
        try:
            workbook = self.get_active_workbook()
            if not workbook:
                excel_app = self._get_excel_app() or self._get_et_app()
                if not excel_app:
                    self.logger.warning("无法启动 Excel 或 WPS 表格")
                    return False
                workbook = excel_app.Workbooks.Add()
            
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
            
            excel_app = self._get_excel_app() or self._get_et_app()
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
