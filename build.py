import sys
api=sys.argv[1] if len(sys.argv)>1 else ''
s=open('src.html',encoding='utf-8').read()
d=open('default.json',encoding='utf-8').read().replace('</','<\\/')
s=s.replace('__DEFAULT__',d).replace("'__API__'",repr(api))
open('index.html','w',encoding='utf-8').write(s)
