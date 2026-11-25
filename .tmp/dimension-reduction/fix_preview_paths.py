from pathlib import Path
p = Path(__file__).with_name('report-preview.html')
s = p.read_text(encoding='utf-8')
s = s.replace(r'src=".tmp\dimension-reduction\docx-preview-media/media/', 'src="docx-preview-media/media/')
p.write_text(s, encoding='utf-8')
