import wx
import pandas as pd
from unidecode import unidecode


class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super(MyFrame, self).__init__(parent, title=title, size=(400, 300))

        self.panel = wx.Panel(self)
        self.data_file_picker = self.create_file_picker("Select Data File:", "data_file_picker")
        self.process_button = wx.Button(self.panel, label="Process", size=(100, -1))

        self.Bind(wx.EVT_BUTTON, self.on_process, self.process_button)

        self.create_layout()
        self.Centre()
        self.Show()

    def create_file_picker(self, label, name):
        file_picker = wx.FilePickerCtrl(
            self.panel, message=label, wildcard="Excel Files (*.xlsx)|*.xlsx")
        return file_picker

    def create_layout(self):
        vbox = wx.BoxSizer(wx.VERTICAL)
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)

        data_label = wx.StaticText(self.panel, label="Data File:")

        hbox1.Add(data_label, flag=wx.RIGHT, border=8)
        hbox1.Add(self.data_file_picker, proportion=1, flag=wx.EXPAND)

        vbox.Add(hbox1, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=10)

        vbox.Add(wx.StaticLine(self.panel), flag=wx.EXPAND | wx.RIGHT | wx.TOP, border=10)

        vbox.Add(wx.StaticLine(self.panel), flag=wx.EXPAND | wx.RIGHT | wx.TOP, border=10)

        vbox.Add(self.process_button, flag=wx.ALIGN_CENTER | wx.TOP, border=10)

        self.panel.SetSizer(vbox)

    def on_process(self, event):
        data_file_path = self.data_file_picker.GetPath()

        if not data_file_path:
            wx.MessageBox("Please select data file.", "Input Error", wx.OK | wx.ICON_ERROR)
            return

        data_df = pd.read_excel(data_file_path)

        keywords = self.show_keywords_dialog()

        if keywords is None:
            return

        results = []

        for index, row in data_df.iterrows():
            cell_value = row["ATP.Texto  L=200"]
            if isinstance(cell_value, str):
                found_keywords = [keyword for keyword in keywords if keyword in cell_value.lower()]
                if found_keywords:
                    results.append([
                        row["PRO.PJ - Protocolo Jurídico"],
                        row["PRO.Número do processo"],
                        cell_value,
                        ", ".join(found_keywords)  # Join the found keywords into a string
                    ])

        if not results:
            wx.MessageBox("No matching records found.", "Results", wx.OK | wx.ICON_INFORMATION)
        else:
            result_df = pd.DataFrame(results, columns=["PRO.PJ - Protocolo Jurídico", "PRO.Número do processo",
                                                       "ATP.Texto L=200", "Palavras-Chave Encontradas"])

            result_file_path = "resultados_da_busca.xlsx"
            result_df.to_excel(result_file_path, index=False)

            wx.MessageBox(f"Matching records saved to '{result_file_path}'.", "Results", wx.OK | wx.ICON_INFORMATION)

    def show_keywords_dialog(self):
        dlg = wx.TextEntryDialog(self, "Enter Keywords (comma-separated):", "Keyword Input", "")
        if dlg.ShowModal() == wx.ID_OK:
            keywords_input = dlg.GetValue()
            keywords = [unidecode(keyword.strip().lower()) for keyword in keywords_input.split(",")]
            dlg.Destroy()
            return keywords
        dlg.Destroy()
        return None


if __name__ == "__main__":
    app = wx.App(False)
    frame = MyFrame(None, "Keyword Processing App")
    app.MainLoop()
