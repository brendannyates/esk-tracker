import os, shutil
from PIL import Image, ImageDraw, ImageFont
os.makedirs('dist/assets', exist_ok=True)
html = open('src/taskpane.html').read()
man = open('src/manifest.xml').read()
assert '`' not in man and '${' not in man
html = html.replace('`__MANIFEST__`', '`' + man.replace('\\','\\\\') + '`')
open('dist/taskpane.html','w').write(html)
open('dist/index.html','w').write('<!doctype html><meta charset="utf-8"><meta http-equiv="refresh" content="0;url=taskpane.html"><title>ESK Tracker</title><a href="taskpane.html">ESK Tracker</a>')
for sz in (16,32,64,80,128):
    S=sz*4; im=Image.new('RGBA',(S,S),(0,0,0,0)); d=ImageDraw.Draw(im)
    d.rounded_rectangle([0,0,S-1,S-1],radius=S//4,fill=(181,83,47,255))
    # simple kit mark: vial outline + checkmark
    w=S*0.28; x0=S*0.36; d.rounded_rectangle([x0,S*0.2,x0+w,S*0.8],radius=S*0.06,outline=(255,253,248,255),width=max(2,S//18))
    d.rectangle([x0-S*0.04,S*0.14,x0+w+S*0.04,S*0.24],fill=(255,253,248,255))
    d.line([(x0+w*0.2,S*0.55),(x0+w*0.45,S*0.68),(x0+w*0.85,S*0.4)],fill=(255,253,248,255),width=max(2,S//16))
    im.resize((sz,sz),Image.LANCZOS).save(f'dist/assets/icon-{sz}.png')
open('dist/commands.html','w').write('<!doctype html><meta charset="utf-8"><title>ESK Tracker commands</title><script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script><script>Office.onReady(function(){});</script>')
print('built')
