import pypyodbc
import csv
import glob
import time
import os
# MS ACCESS DB CONNECTION
pypyodbc.lowercase = False

def getTablesName(filename):
    conn = pypyodbc.connect(
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
    r"Dbq=%s;" % filename)
    cur = conn.cursor()
    tablename = None

    tablelist = []
    for i in cur.tables():
        if i[3] == 'TABLE':
            tablelist.append(i[2])
    cur.close()
    conn.close()
    return tablelist

    
def mdbtotxt(f,filename,tname):
    conn = pypyodbc.connect(
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
    r"Dbq=%s;" % filename)
    cur = conn.cursor()
    cur.execute('select * from %s' % tname)
    for row in cur.fetchall() :
        f.write(str(row)+'\n')
    cur.close()
    conn.close()
    time.sleep(1)

def run(text):    
    for filename in glob.glob(text):
        filename = filename.replace('\\','/')
        print(filename)
        tablelist = getTablesName(filename)
        curDate = time.strftime("-%Y-%m-%d--%H_%M_%S", time.localtime())
        newfile = os.path.splitext(os.path.basename(filename))[0] + curDate + '.txt'
        print(newfile)
        f = open(newfile,'w', encoding='utf-8')
        for tname in tablelist:
            print(tname)
            mdbtotxt(f,filename,tname)
            f.flush()
        f.close()
if __name__ == '__main__':
    run(r'E:\FeigeDownload\*.mdb')
    
    print('over ... ')


            
            