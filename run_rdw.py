from pathlib import Path
from tkinter import Tk, filedialog, messagebox
from rdw.engine import process

root = Tk()
root.withdraw()
messagebox.showinfo('RDW v0.1', 'Select a Word document to de-identify.')
source = filedialog.askopenfilename(title='Select Word document', filetypes=[('Word documents', '*.docx')])
if source:
    dictionary = filedialog.askopenfilename(title='Select replacement dictionary CSV', filetypes=[('CSV files', '*.csv')])
    if dictionary:
        output = filedialog.askdirectory(title='Select output folder')
        if output:
            try:
                result, log, manifest = process(source, dictionary, output)
                messagebox.showinfo('RDW complete', f'Created: {result.name}\nAudit: {log.name}\nReplacements: {manifest["replacement_events"]}\n\nManually review output before sharing.')
            except Exception as exc:
                messagebox.showerror('RDW error', str(exc))
root.destroy()
