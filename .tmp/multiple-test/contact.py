from PIL import Image
from pathlib import Path
fs=sorted(Path('.tmp/multiple-test/final-pages').glob('page-*.png'))
for i in range(0,len(fs),2):
    ims=[Image.open(f).convert('RGB') for f in fs[i:i+2]]
    sheet=Image.new('RGB',(sum(x.width for x in ims),max(x.height for x in ims)), '#dddddd')
    offset=0
    for im in ims:
        sheet.paste(im,(offset,0))
        offset+=im.width
    sheet.save('.tmp/multiple-test/final-pages/pair-'+str(i//2+1)+'.png')
