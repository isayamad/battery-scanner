from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from datetime import datetime
import os

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

class BatteryApp(App):
    def build(self):
        self.excel_path = os.path.join(os.path.expanduser("~"), "BatteryLog.xlsx")
        self.init_excel()
        
        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        layout.add_widget(Label(text="🔋 Battery Logger", font_size=24, size_hint_y=0.1))
        
        self.status = Label(text="Ready", size_hint_y=0.1)
        layout.add_widget(self.status)
        
        btn_manual = Button(text="📝 Enter Battery ID", size_hint_y=0.15, font_size=18)
        btn_manual.bind(on_press=self.manual_entry)
        layout.add_widget(btn_manual)
        
        btn_view = Button(text="📋 View Log", size_hint_y=0.15, font_size=18)
        btn_view.bind(on_press=self.view_log)
        layout.add_widget(btn_view)
        
        self.count_label = Label(text=self.get_count(), size_hint_y=0.1)
        layout.add_widget(self.count_label)
        
        return layout
    
    def init_excel(self):
        if not EXCEL_OK:
            return
        if not os.path.exists(self.excel_path):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws['A1'] = "Battery ID"
            ws['B1'] = "Date"
            ws['C1'] = "Time"
            wb.save(self.excel_path)
    
    def get_count(self):
        if not EXCEL_OK or not os.path.exists(self.excel_path):
            return "Total: 0"
        wb = openpyxl.load_workbook(self.excel_path, read_only=True)
        count = max(0, wb.active.max_row - 1)
        wb.close()
        return f"Total: {count}"
    
    def manual_entry(self, instance):
        content = BoxLayout(orientation='vertical', padding=10, spacing=10)
        content.add_widget(Label(text="Enter Battery ID:", size_hint_y=0.2))
        
        text_input = TextInput(hint_text="e.g. LEADPOWER 12345", multiline=False)
        content.add_widget(text_input)
        
        btn_box = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=0.3)
        submit = Button(text="Submit")
        cancel = Button(text="Cancel")
        btn_box.add_widget(submit)
        btn_box.add_widget(cancel)
        content.add_widget(btn_box)
        
        popup = Popup(title="Manual Entry", content=content, size_hint=(0.9, 0.4))
        
        def do_submit(btn):
            battery_id = text_input.text.strip()
            if battery_id:
                self.log_battery(battery_id)
                popup.dismiss()
        
        submit.bind(on_press=do_submit)
        cancel.bind(on_press=popup.dismiss)
        popup.open()
    
    def log_battery(self, battery_id):
        if not EXCEL_OK:
            self.status.text = "Excel not available"
            return
        
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb.active
        
        # Check for duplicate
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == battery_id:
                self.status.text = f"Duplicate: {battery_id}"
                wb.close()
                return
        
        # Add new entry
        now = datetime.now()
        row = ws.max_row + 1
        ws.cell(row=row, column=1, value=battery_id)
        ws.cell(row=row, column=2, value=now.strftime("%Y-%m-%d"))
        ws.cell(row=row, column=3, value=now.strftime("%H:%M:%S"))
        wb.save(self.excel_path)
        
        self.status.text = f"Added: {battery_id}"
        self.count_label.text = self.get_count()
    
    def view_log(self, instance):
        if not EXCEL_OK or not os.path.exists(self.excel_path):
            self.status.text = "No log file"
            return
        
        content = ScrollView()
        log_layout = BoxLayout(orientation='vertical', size_hint_y=None, spacing=2)
        log_layout.bind(minimum_height=log_layout.setter('height'))
        
        wb = openpyxl.load_workbook(self.excel_path, read_only=True)
        for row in wb.active.iter_rows(min_row=2, values_only=True):
            if row[0]:
                label = Label(text=f"{row[0]} - {row[1]} {row[2]}", 
                             size_hint_y=None, height=30, font_size=12)
                log_layout.add_widget(label)
        wb.close()
        
        content.add_widget(log_layout)
        popup = Popup(title="Battery Log", content=content, size_hint=(0.95, 0.8))
        popup.open()

if __name__ == "__main__":
    BatteryApp().run()
