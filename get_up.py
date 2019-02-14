#encoding:utf-8
import codecs

def get_up( up_nb ):
    discip = set()
    with codecs.open( "up_content/%09d.txt" % up_nb, 'r') as f:
        lines = [ l.replace('\n','').replace('\r','').split('\t') for l in f.readlines() ]
        for l in lines:
          if len(l) > 1:
            if l[1] != "": discip.add( l[0].strip() )

    return sorted(list(discip))

