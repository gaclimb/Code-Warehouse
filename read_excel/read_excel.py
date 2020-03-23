import xlrd
import sqlite3

# 读取excel内容并整理成元组的形式
def read_excel():
    data = []
    book = xlrd.open_workbook(file_name)
    sheet = book.sheet_by_index(0)
    rows = sheet.nrows
    for i in range(1, rows):
        row = sheet.row_values(i)
        data.append(tuple(row))
    return data

# 将数据写入到数据库
def db(row):
    conn = sqlite3.connect(db_file)
    cur = conn.cursor()
    for i in row:
        sql = 'insert into bg(id,dw,lx,dh,ly,xt,nr,cs,rq,sj,zd,fw,ry,yw,sh) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
        cur.execute(sql, i)
        conn.commit()
    cur.close()
    conn.close()


if __name__ == '__main__':
    file_name = '变更.xlsx'
    db_file = 'DB.db'
    row = read_excel()
    db(row)
