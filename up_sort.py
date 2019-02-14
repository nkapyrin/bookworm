#encoding:utf-8

import codecs

with codecs.open( "up_list.txt", 'r', encoding="utf-8" ) as upf:
  up_lines = [  l.replace('\n','').replace('\r','').strip().strip(';') for l in upf.readlines() ]
  up_lst = [ l.split(';') for l in up_lines ]
  print u' '.join(up_lst[0])
  up_srt = sorted( up_lst[1:], key=lambda e:(e[1],e[4]), reverse=True )
  up_dict = [ dict(zip(up_lst[0],up)) for up in up_srt ]

for up in up_dict:
  print '|'.join( [up[u"Год"],up[u"Код"], up[u"Шифр профиля/специализации"], up[u"Профиль/Специализация"],up[u"Уровень обучения"]])
  #print up[6]
