from pathlib import Path
import win32com.client
import sys
ROOT=Path(__file__).resolve().parents[2]
app=win32com.client.DispatchEx('Word.Application')
app.Visible=False
app.DisplayAlerts=0
doc=None
try:
    print('Word instance created',flush=True)
    input_path=Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/'作业报告/多重检验校正报告_字体公式修正版.docx'
    output_path=Path(sys.argv[2]) if len(sys.argv)>2 else ROOT/'.tmp/multiple-test/report-font-fixed.pdf'
    doc=app.Documents.Open(str(input_path.resolve()),ConfirmConversions=False,ReadOnly=True,AddToRecentFiles=False,Visible=False,OpenAndRepair=False)
    print('Document opened',flush=True)
    doc.Repaginate()
    print('Pages:',doc.ComputeStatistics(2),flush=True)
    doc.ExportAsFixedFormat(str(output_path.resolve()),17)
    print('Exported',flush=True)
finally:
    if doc is not None:doc.Close(0)
    app.Quit(0)
