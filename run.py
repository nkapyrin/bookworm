#encoding:utf-8
import os, docx, codecs, re
from os.path import isfile, join
from shorten_bib_entry import shorten

#from get_up import get_up

bib = set()
dict_by_discipline = {}

up_field_names = [u"﻿Наименование",u"Профиль/Специализация",u"Шифр профиля/специализации",u"Площадка",u"Год",\
                  u"Форма обучения",u"Уровень обучения",u"Для иностранных групп",u"Выпускающая кафедра",\
                  u"Статус разработки",u"Утверждено проректором",u"Приведен к шаблону",u"Скан прикреплен",\
                  u"Дата отправки на утверждение",u"Дата утверждения",u"Код",u"Нет набора",u"Для лицензирования" ]
rpd_field_names = [ u"Наименование", u"Дисциплина", u"Учебный план", u"Шифр", u"Выпускающая кафедра",\
                    u"Обеспечивающая кафедра",u"Код", u"Год поступления", u"Статус разработки", \
                    u"Дата отправки на утверждение",u"Дата утверждения",u"Утверждено проректором"]
rpd_by_plan = {}

# Загрузить дисциплины для каждого УП
#discip_by_plan = {}
#with codecs.open( "up_list.txt", 'r', encoding="utf-8" ) as upf:
#    up_lines = [  l.replace('\n','').replace('\r','').strip().strip(';').split() for l in upf.readlines() ]
#    up_lst = [ l.split(';') for l in up_lines[1:] ]
#    up_dict = [ dict(zip(up_field_names, l)) for l in sorted( up_lst, key=lambda e:(e[1],e[4]), reverse=True ) ]
#    for up in up_dict:
#       print int(u"Код")
#       discip_by_plan[ int(u"Код") ] = get_up( int(u"Код"))

year_re = re.compile( u'2[0-9]{3}/[0-9]{2}' )
rpd_re = re.compile( u', [0-9]{9}$' )
#spec_re = [ re.compile( u'[0-9]{2}.[0-9]{2}.[0-9]{2}' ), re.compile( u'[0-9]{6}' ) ]



up_info_keys = [ u"Наименование", u"Профиль/Специализация", u"Шифр профиля/специализации", u"Площадка", u"Год", u"Форма обучения", u"Уровень обучения", u"Для иностранных групп", u"Выпускающая кафедра", u"Статус разработки", u"Утверждено проректором", u"Приведен к шаблону", u"Скан прикреплен", u"Дата отправки на утверждение", u"Дата утверждения", u"Код", u"Нет набора", u"Для лицензирования" ]
up_info_by_id = {}
with codecs.open( u'Список УП.txt', encoding='utf-8') as f:
    lines = [ l.strip(' ').replace('\n','').replace('\r','' ) for l in f.readlines() ]
    d = [dict( zip(up_info_keys, l.split('\t') )) for l in lines[1:]]
    for dd in d: up_info_by_id[ int(dd[u"Код"]) ] = dd

#
# Из списка РПД по дисциплинам, загружаем все номера РПД и ЛУ
# в  [up_dis_code]  :::  код УП --> название дисциплины --> код РПД
#
up_dis_code = {}

rpd_name_by_id = {}
with codecs.open( u"РПД по УП 305_2018_11_18.txt", 'r', encoding="utf-8" ) as upf:
    lines = [ l.replace('\n','').replace('\r','').strip().strip(';') for l in upf.readlines() ]
    lst = [ l.split('\t') for l in lines[9:] ]
    last_up = None
    last_year = None
    last_discip = None
    for l in lst:
        first = l[0]
        if year_re.match( first ):
            #print ">>> год <<<", first
            last_year = first
            up_dis_code[ last_year ] = dict()
        elif u"УП" in first[0:2]:
            #print u">>> УП <<<", l[4]
            last_up = l[4]
            up_dis_code[ last_year ][ last_up ] = dict()
        elif rpd_re.search( first ):
            #if u"Лист переутверждения" in first: print u">>>  ЛУ <<<", first.split(', ')[-1]
            #else: print u">>> РПД <<<", first.split(', ')[-1]
            #if last_discip: 
            if u"Лист переутверждения" in first:
              up_dis_code[ last_year ][ last_up ][ last_discip ] = u"ЛП " + first.split(', ')[-1]
              rpd_name_by_id[ int(first.split(', ')[-1]) ] = u'|'.join(l)
            else:
              up_dis_code[ last_year ][ last_up ][ last_discip ] = u"РПД " + first.split(', ')[-1]
              rpd_name_by_id[ int(first.split(', ')[-1]) ] = u'|'.join(l)
            #print up_dis_code[ last_year ][ last_up ][ last_discip ]
        else:
            #if last_up: print '>>>  D  <<<', first
            #else: # Такого случая не должно быть
        	#    print first
        	#    last_up = None
            last_discip = first
            up_dis_code[ last_year ][ last_up ][ last_discip ] = ""


print '========================================'

    #srt = sorted( up_lst, key=lambda e:(e[1],e[4]), reverse=True )


# !!!!!!!!!!!!!!!
# Отсортировать по годам
udc = {}
total_rpd_list = []
# Пройти все года
for year in sorted( up_dis_code.keys() ):
	print '____________________________ year', year
	# Для каждого Кода УП каждого года, найти его Шифр Специализации
	for up_id in up_dis_code[year].keys():
		# Этого УП нет в списке
		if int(up_id) not in up_info_by_id.keys(): print '! ' + up_id; continue;
		up_short = up_info_by_id[ int( up_id )][u"Шифр профиля/специализации"]
		# Отсеять не интересующие нас специальности
		#if up_short not in []: print up_short; continue
		#if up_short not in [u"24.03.02.Б14", u"24.05.05.С5", u"24.05.06.С13", u"24.05.06.Б14", u"24.05.06.С11", u"24.05.06.С12"]: print 'up', up_short, 'not in scope'; continue
		#if up_short in [u"24.03.02.Б14", u"24.05.05.С5", u"24.05.06.С13", u"24.05.06.Б14", u"24.05.06.С11", u"24.05.06.С12", u'161101.С13', u'161101.С12', u'161101.С11', u'161400.С5']: print 'up', up_short, 'not in scope'; continue
		#else: print 'up', up_short, 'has', len(up_dis_code[year][up_id]), 'RPD'
		print 'up', up_short, 'has', len(up_dis_code[year][up_id]), 'RPD'
		if up_short not in udc.keys(): udc[ up_short ] = dict()
		for discip_name in up_dis_code[year][up_id].keys():
			#print up_dis_code[ year ][ up_id ][ discip_name ]
			if discip_name not in udc[ up_short ].keys() and u"РПД" in up_dis_code[ year ][ up_id ][ discip_name ]:
				print discip_name
				udc[ up_short ][ discip_name ] = { 'rpd_id':up_dis_code[ year ][ up_id ][ discip_name ] }
				total_rpd_list.append( udc[ up_short ][ discip_name ]['rpd_id'] )
		#print len( total_rpd_list )

# !!!!!!!!!!!!!!!

#for k1 in udc.keys():
#	print '  ' + k1
#	for k2 in udc[k1].keys():
#		print '     ' + k2 + ' : ' + udc[k1][k2]

missing_rpd = []
for k1 in udc.keys():
	print '---' + k1
	for k2 in sorted(udc[k1].keys()):
		#print k2
		rpd_nb = udc[k1][k2]['rpd_id'].strip(u'РПД ')
		if not isfile( join(u'rpd_txt_files', u'rpd'+rpd_nb+u'.txt') ):
			missing_rpd.append( rpd_nb )
			#print '   ', join(u'rpd_txt_files', u'rpd'+rpd_nb+u'.txt'), "missing"
			continue;
		#else:
		#	print '   ', join(u'rpd_txt_files', u'rpd'+rpd_nb+u'.txt'), "OK"
		# Прочитать файл РПД и найти всю нужную информацию
		#with codecs.open( join(u'rpd_txt_files',fnRPD),'r', encoding="utf-8" ) as rpdf:

print 'total_rpd_list', len(total_rpd_list)
print 'missing_rpd', len(missing_rpd)
print '\n'.join( str(s) + ' ' + rpd_name_by_id[s] for s in sorted( [int(i) for i in missing_rpd] ) )




for k1 in udc.keys():
	print '  ' + k1
	for ik2,k2 in enumerate( sorted(udc[k1].keys())):
		rpd_nb = udc[k1][k2]['rpd_id'].strip(u'РПД ')
		udc[k1][k2][u"Основная литература"] = []
		udc[k1][k2][u"Дополнительная литература"] = []
		udc[k1][k2]["code"] = rpd_nb

		if not isfile( join(u'rpd_txt_files', u'rpd'+rpd_nb+u'.txt') ):
			print '   ', join(u'rpd_txt_files', u'rpd'+rpd_nb+u'.txt'), "missing"; continue;

		# Прочитать файл РПД и найти всю нужную информацию
		with codecs.open( join(u'rpd_txt_files', u'rpd'+rpd_nb+u'.txt'), 'r', encoding="utf-8" ) as rpdf:
			rpd_lines = [  l.replace('\n','').replace('\r','').replace(u'¬','').strip().strip(';') for l in rpdf.readlines() ]
			cur_discip_nb = ""; cur_discip_name = ""; writing_main_bibliography = 0; writing_extra_bibliography = 0; writing_ecatalog_bib = 0; #next_is_discipline_name = 0;
			# Просмотреть все строчки файла РПД
			for rpdl in rpd_lines:
				if len(rpdl.strip()) == 0: continue
				elif u"основная литература:" in rpdl: writing_extra_bibliography, writing_main_bibliography = 0, 1;
				elif u"дополнительная литература:" in rpdl: writing_extra_bibliography, writing_main_bibliography = 1, 0;
				elif u"Литература из электронного каталога" in rpdl or u"Дополнительная литература" in rpdl: continue;
				elif l.strip(' ') == "" or u"ПЕРЕЧЕНЬ РЕСУРСОВ ИНФОРМАЦИОННО" in rpdl or u'МАТЕРИАЛЬНО-ТЕХНИЧЕСКОЕ ОБЕСПЕЧЕНИЕ ПРАКТИКИ' in rpdl or u"Интернет-ресурсы" in rpdl or u"Интернет-сайты" in rpdl or u"Периодические издания" in rpdl:
					writing_extra_bibliography, writing_main_bibliography = 0, 0;
				elif writing_main_bibliography == 1 or writing_extra_bibliography == 1:
					txt = rpdl.strip().replace(u"•",u"")
					if txt[0] in u'1234567890':
						if txt.find(u'.') < 4 : separator = u'.'
						elif txt.find(u')') < 4: separator = u')' 
						else: separator = u' '
						txt = txt[ txt.find(separator)+1: ].strip()
					if len(txt.strip()) > 0:
						if writing_main_bibliography==1:  udc[k1][k2][u"Основная литература"].append( txt )
						elif writing_extra_bibliography==1: udc[k1][k2][u"Дополнительная литература"].append( txt )
		print ik2+1, k2, rpd_nb, len(udc[k1][k2][u"Основная литература"]), len(udc[k1][k2][u"Дополнительная литература"])

#exit()

# Сохранить на диск
with codecs.open( "books.txt", 'w' ) as fb:
	with codecs.open( "records.txt", 'w' ) as fr:
		for u in udc.keys():
			#f.write( (u"\n" + up_id + u"\n").encode('utf-8') )
			for d in udc[u].keys():
				for b in ( udc[u][d][u"Основная литература"] + udc[u][d][u"Дополнительная литература"]):
					fb.write( (b + u"\n").encode('utf-8') )
					if len( shorten(b) ) > 0: fr.write( (shorten(b) + '\n').encode('utf-8') )

# Сохранить на диск
#with codecs.open( "records.txt", 'w' ) as f:
#  for u in udc.keys():
#    for d in udc[u].keys():
#      for b in (udc[u][d][u"Основная литература"] + udc[u][d][u"Дополнительная литература"]):
#          if len( shorten(b) ) > 0: f.write( (shorten(b) + '\n').encode('utf-8') )



# Загрузить данные из библиотеки
os.system( 'python query_library.py' )



#print '---'
#for up_id in rpd_by_plan.keys(): print up_id
#print '---'



books_info = {}
keys_lst = [ u"Код", u"Тип", u"Запись", u"Авторы", u"Экземпляры", u"Шифры", u"Ссылка" ]
with codecs.open( "books_records_out.txt", 'r' ) as f:
    lines = [  l.replace('\n','').replace('\r','').decode("utf-8") for l in f.readlines() ]
    lines_lst = [ l.split('|') for l in lines ]
    for fields in lines_lst:
        books_info[ fields[0] ] = dict( zip(keys_lst, fields))
        ex = books_info[ fields[0] ][u"Экземпляры"]
        books_info[ fields[0] ][u"Всего"] = ex[ ex.find(u"Всего")+7:ex.find(u",") ]


list_books_in_lib = []
list_books_not_in_lib = []

for u in udc.keys():
	template = docx.Document('report_template.docx');
	print u
	for tb in template.tables:
		if tb.rows[0].cells[0].paragraphs[0].text == u"#TABLE_BY_PLAN_INSERT_PLAN_ID_AND_NAME#":
			tb.rows[0].cells[0].paragraphs[0].text = u # up["header"]
			src_row = tb.rows[-1]; src_cells = src_row.cells;
			for iid,d in enumerate( sorted(udc[u].keys()) ):
				print iid, d
				first_row_cells, last_row_cells = None, None
				if len( udc[u][d][u"Основная литература"] + udc[u][d][u"Дополнительная литература"]) == 0:
					tb.add_row()
					first_row_cells = tb.rows[-1].cells
					first_row_cells[0].text = str(iid+1)
					first_row_cells[2].text = d # + "\n" + rpd[u"Код"]
				else:
					for ib,b in enumerate(udc[u][d][u"Основная литература"] + udc[u][d][u"Дополнительная литература"]):
						tb.add_row()
						#print '>>>', b
						if ib==0: first_row_cells = tb.rows[-1].cells
						last_row_cells = tb.rows[-1].cells
                  
						short = shorten( b )
						last_row_cells[3].text = b

						# Если текущая запись есть в перечне, подставим количество книг
						if short in books_info.keys():
							last_row_cells[4].text = books_info[short][u"Всего"]
							if len( books_info[short][u"Ссылка"] ) > 0: last_row_cells[5].text = u"✓"
							list_books_in_lib.append( short )
							# Если такой записи нет, попробуем подставить запись, где книга издана в другом году
						elif '$' in short and short.split('$')[0] in books_info.keys():
							short = short.split('$')[0]
							last_row_cells[4].text = books_info[short][u"Всего"] + u'*'
							if len( books_info[short][u"Ссылка"] ) > 0: last_row_cells[5].text = u"✓"
							list_books_in_lib.append( short )
						else: list_books_not_in_lib.append( short )
					#
					for iic,c in enumerate(tb.rows[-1].cells): c.paragraphs[0].style = src_cells[iic].paragraphs[0].style

					# Вставить название предмета
					if first_row_cells and last_row_cells:
						A = first_row_cells[0].merge( last_row_cells[0] ); A.text = str(iid+1)
						B = first_row_cells[2].merge( last_row_cells[2] ); B.text = d #+ "\n" + udc[u][d]["code"]
						first_row_cells[1].merge( last_row_cells[1] );
						first_row_cells[6].merge( last_row_cells[6] );

		tb._tbl.remove(src_row._tr)
	template.save( u"out/Литература по предметам %s.docx" % u )
	print 'saved > ' + u"out/Литература по предметам %s.docx" % u
	print '\n'

#print "missing books"
#for b in list_books_not_in_lib: print b
print "missing books (%d, %d unique):" % ( len(list_books_not_in_lib), len( list(set(list_books_not_in_lib))) )
print "OK books (%d):" % len(list_books_in_lib)









exit()


up_info = {}
with codecs.open( "up_list.txt", 'r', encoding="utf-8" ) as upf:
    up_lines = [  l.replace('\n','').replace('\r','').strip().strip(';') for l in upf.readlines() ]
    up_lst = [ l.split(';') for l in up_lines[1:] ]
    up_srt = sorted( up_lst, key=lambda e:(e[1],e[4]), reverse=True )

    for up_fields in up_srt: #in up_lines[1:]:
        #print len(up_line.split(';'))
        up_tmp = dict( zip(up_field_names, [s.replace('\n','').replace('\r','').strip() for s in up_fields] )  )
        file_nb = str(int( up_tmp[u"Код"] ))
        # Пройтись по всем РПД этого УП
        if not isfile( join(u'lists_disciplines',file_nb + ".csv") ):
           print join(u'lists_disciplines',file_nb + ".csv"), "missing"
           continue;
        else: print join(u'lists_disciplines',file_nb+".csv"), "OK"
        #up_id =  up[u"Шифр профиля/специализации"] + u" " + up[u"Год"].replace('/','-')
        up_id =  up_tmp[u"Шифр профиля/специализации"]

        if up_id not in up_info.keys(): up_info[up_id] = []
        up_info[up_id].append( up_tmp )


print ''
for k in up_info.keys(): print k, [up[u'Год'] for up in sorted(up_info[k], key=lambda e:e[u"Год"] )]
print ''


for k in up_info.keys():
    print '\n>>>', k
    for up in sorted(up_info[k], key=lambda e:e[u"Год"] ):
    	for k2 in up_dis_code[ up[u"Год"] ][ up[u"Код"] ]:
            print up_dis_code[ up[u"Год"] ][ up[u"Код"] ][k2]

            if u"РПД" not in up_dis_code[ up[u"Год"] ][ up[u"Код"] ][k2]: continue

            file_nb = str(int( up_tmp[u"Код"] ))





            #up_id =  up[u"Шифр профиля/специализации"] + u" " + up[u"Год"].replace('/','-')
            up_id =  up[u"Шифр профиля/специализации"]
            if up_id not in rpd_by_plan.keys():
                rpd_by_plan[ up_id ] = up
                rpd_by_plan[ up_id ]["rpd"] = {}
                rpd_by_plan[ up_id ]["id"] = up_id
                rpd_by_plan[ up_id ]["header"] = u"Специализация " + up[u"Шифр профиля/специализации"] + " " + up[u"Профиль/Специализация"]
            if up_id not in rpd_by_plan.keys(): rpd_by_plan[up_id] = {}




            ######################
            # TODO: считать все РПД дисциплин, практик + все ЛУ
            ######################




            # Открываем РПД
            for folder in [ 'lists_disciplines', 'lists_internship' ]:
                filename = ''
                if isfile( join( folder, file_nb+".csv" )): filename = file_nb + ".csv"
                elif isfile( join( folder, file_nb+".txt" )): filename = file_nb + ".txt"
                else:
                    print "No RPD" + file_nb
                    continue
            
                with codecs.open( join( folder, filename), 'r', encoding="utf-8" ) as f:
                    if '.csv' in filename: lines = [  l.replace('\n','').replace('\r','').strip().strip(';') for l in f.readlines() ]
                    else: lines = [  l.replace('\n','').replace('\r','').strip().strip('\t') for l in f.readlines() ]
                    for line in lines[1:]:
                        if '.csv' in filename: rpd_tmp = dict( zip(rpd_field_names, line.split(';'))  )
                        else: rpd_tmp = dict( zip(rpd_field_names, line.split('\t'))  )
                        #if u"305" not in rpd[u"Выпускающая кафедра"]: continue    # если нужно исключить РПД других кафедр 
                        rpd_tmp[u"Основная литература"] = []
                        rpd_tmp[u"Дополнительная литература"] = []
                
                        fnRPD = u"rpd"+rpd_tmp[u"Код"]+".txt"
                        idRPD = rpd_tmp[u"Дисциплина"]

                        if idRPD not in rpd_by_plan[ up_id ]["rpd"].keys(): rpd_by_plan[ up_id ]["rpd"][ idRPD ] = rpd_tmp
                        else: continue # Не будем обрабатывать более старые РПД (УП отсортированы по годам)

                        # Если такого файла нет, ничего не поделаешь
                        if not isfile( join(u'rpd_txt_files',fnRPD) ):
                            print '   ', join(u'rpd_txt_files',fnRPD), "missing"
                            continue;
                        # Прочитать файл РПД и найти всю нужную информацию
                        with codecs.open( join(u'rpd_txt_files',fnRPD),'r', encoding="utf-8" ) as rpdf:
                            rpd_lines = [  l.replace('\n','').replace('\r','').replace(u'¬','').strip().strip(';') for l in rpdf.readlines() ]
                            cur_discip_nb = ""; cur_discip_name = "";
                            writing_main_bibliography = 0; writing_extra_bibliography = 0; writing_ecatalog_bib = 0; #next_is_discipline_name = 0;
                            # Просмотреть все строчки файла РПД
                            for rpdl in rpd_lines:
                                if len(rpdl.strip()) == 0: continue
                                elif u"а)основная литература:" in rpdl:
                                    writing_extra_bibliography = 0; writing_main_bibliography = 1;
                                elif u"б)дополнительная литература:" in rpdl:
                                    writing_extra_bibliography = 1; writing_main_bibliography = 0;
                                elif u"Литература из электронного каталога" in rpdl or u"Дополнительная литература" in rpdl: continue;
                                elif l.strip(' ') == "" or u"ПЕРЕЧЕНЬ РЕСУРСОВ ИНФОРМАЦИОННО" in rpdl or u'МАТЕРИАЛЬНО-ТЕХНИЧЕСКОЕ ОБЕСПЕЧЕНИЕ ПРАКТИКИ' in rpdl or u"Интернет-ресурсы" in rpdl or u"Интернет-сайты" in rpdl or u"Периодические издания" in rpdl:
                                    writing_extra_bibliography = 0; writing_main_bibliography = 0;
                                elif writing_main_bibliography == 1:
                                    txt = rpdl.strip()
                                    if txt[0] in u'1234567890':
                                      separator = u'.' if txt.find(u'.') < 4 else u')'
                                      txt = txt[ txt.find(separator)+1: ].strip()
                                    if len(txt.strip()) > 0: rpd_by_plan[ up_id ]["rpd"][idRPD][u"Основная литература"].append( txt )
                                elif writing_extra_bibliography==1:
                                    txt = rpdl.strip()
                                    if txt[0] in u'1234567890':
                                        separator = u'.' if txt.find(u'.') < 4 else u')'
                                        txt = txt[ txt.find(separator)+1: ].strip()
                                        #txt = txt[ txt.find(u'.')+1: ].strip()
                                    if len(txt.strip()) > 0: rpd_by_plan[ up_id ]["rpd"][idRPD][u"Дополнительная литература"].append( txt )
                    # Добавить дисциплины, заявленные в учебном плане, но не встреченные в РПД (отладка)
                    #for up_dis in discip_by_plan[int( up_tmp[u"Код"] )]:
                    #	if up_dis not in rpd_by_plan[ up_id ]["rpd"].keys(): rpd_by_plan[ up_id ]["rpd"][up_dis] = {}


# Сохранить на диск
with codecs.open( "books.txt", 'w' ) as f:
  for up_id in rpd_by_plan.keys():
    #f.write( (u"\n" + up_id + u"\n").encode('utf-8') )
    for rpd_id in rpd_by_plan[up_id]['rpd']:
      #f.write( (rpd_id + '\n').encode('utf-8') )
      rpd = rpd_by_plan[up_id]['rpd'][rpd_id]
      for b in (rpd[u"Основная литература"]+rpd[u"Дополнительная литература"]):
          f.write( (b + u"\n").encode('utf-8') )

# Сохранить на диск
with codecs.open( "records.txt", 'w' ) as f:
  for up_id in rpd_by_plan.keys():
    for rpd_id in rpd_by_plan[up_id]['rpd']:
      rpd = rpd_by_plan[up_id]['rpd'][rpd_id]
      #f.write( (rpd_id + " " + rpd[u"Код"] + '\n').encode('utf-8') )
      for b in (rpd[u"Основная литература"] + rpd[u"Дополнительная литература"]):
          if len( shorten(b) ) > 0: f.write( (shorten(b) + '\n').encode('utf-8') )

# Загрузить данные из библиотеки
os.system( 'python query_library.py' )


#print '---'
#for up_id in rpd_by_plan.keys(): print up_id
#print '---'


books_info = {}
keys_lst = [ u"Код", u"Тип", u"Запись", u"Авторы", u"Экземпляры", u"Шифры", u"Ссылка" ]
with codecs.open( "books_records_out.txt", 'r' ) as f:
    lines = [  l.replace('\n','').replace('\r','').decode("utf-8") for l in f.readlines() ]
    lines_lst = [ l.split('|') for l in lines ]
    for fields in lines_lst:
        books_info[ fields[0] ] = dict( zip(keys_lst, fields))
        ex = books_info[ fields[0] ][u"Экземпляры"]
        books_info[ fields[0] ][u"Всего"] = ex[ ex.find(u"Всего")+7:ex.find(u",") ]


list_books_in_lib = []
list_books_not_in_lib = []

for up_id in rpd_by_plan.keys():
    up = rpd_by_plan[up_id]
    file_nb = str(int( up[u"Код"] ))
    template = docx.Document('report_template.docx');
    print up_id
    for tb in template.tables:
        if tb.rows[0].cells[0].paragraphs[0].text == u"#TABLE_BY_PLAN_INSERT_PLAN_ID_AND_NAME#":
            tb.rows[0].cells[0].paragraphs[0].text = up["header"]
            src_row = tb.rows[-1]; src_cells = src_row.cells;
            for ii,idRPD in enumerate( sorted(rpd_by_plan[ up_id ]["rpd"].keys()) ):
                #print idRPD
                rpd = rpd_by_plan[ up_id ]["rpd"][idRPD]
                first_row_cells, last_row_cells = None, None
                for ib,b in enumerate(rpd[u"Основная литература"] + rpd[u"Дополнительная литература"]):
                  tb.add_row()
                  #print '>>>', b
                  if ib==0: first_row_cells = tb.rows[-1].cells
                  last_row_cells = tb.rows[-1].cells
                  
                  short = shorten( b )
                  #if short in books_info.keys():
                  #     last_row_cells[3].text = b + "\n=======\n" + books_info[short][u"Запись"] + "\n///\n" + short
                  #elif '$' in short and short.split('$')[0] in books_info.keys():
                  #     last_row_cells[3].text = b + "\n=======\n" + books_info[short.split('$')[0]][u"Запись"] + "\n///\n" + short.split('$')[0]
                  #else: last_row_cells[3].text = b + short
                  last_row_cells[3].text = b

                  # Если текущая запись есть в перечне, подставим количество книг
                  if short in books_info.keys():
                      last_row_cells[4].text = books_info[short][u"Всего"]
                      if len( books_info[short][u"Ссылка"] ) > 0: last_row_cells[5].text = u"✓"
                      list_books_in_lib.append( short )
                  # Если такой записи нет, попробуем подставить запись, где книга издана в другом году
                  elif '$' in short and short.split('$')[0] in books_info.keys():
                      short = short.split('$')[0]
                      last_row_cells[4].text = books_info[short][u"Всего"] + u'*'
                      if len( books_info[short][u"Ссылка"] ) > 0: last_row_cells[5].text = u"✓"
                      list_books_in_lib.append( short )
                  else: list_books_not_in_lib.append( short )
                  #
                  for iic,c in enumerate(tb.rows[-1].cells): c.paragraphs[0].style = src_cells[iic].paragraphs[0].style

                # Вставить название предмета
                if first_row_cells and last_row_cells:
                  A = first_row_cells[0].merge( last_row_cells[0] ); A.text = str(ii+1)
                  B = first_row_cells[2].merge( last_row_cells[2] ); B.text = rpd[ u"Дисциплина" ] # + "\n" + rpd[u"Код"]
                  first_row_cells[1].merge( last_row_cells[1] );
                  first_row_cells[6].merge( last_row_cells[6] );

            print '\n'
            tb._tbl.remove(src_row._tr)
    template.save( u"out/Литература по предметам %s.docx" % up_id )

#print "missing books"
#for b in list_books_not_in_lib: print b
print "missing books (%d, %d unique):" % ( len(list_books_not_in_lib), len( list(set(list_books_not_in_lib))) )
print "OK books (%d):" % len(list_books_in_lib)
