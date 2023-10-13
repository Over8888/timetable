# timetable<br>
#大框架<br>
import cx_Oracle<br>

from docx import Document<br>

#数据库连接信息<br>
username = ''<br>
password = ''<br>
host = ''<br>
port = ''<br>
#service_name = '.JSKB'  # 请替换为你的数据库服务名称<br>

#连接到Oracle数据库<br>
dsn = cx_Oracle.makedsn('', '', service_name='orcl') <br>
connection = cx_Oracle.connect(user='', password='', dsn=dsn)<br>
#创建游标<br>
cursor = connection.cursor()<br>

#执行查询<br>
query = "SELECT KCM,ZC,JASDM FROM .JSKB"<br>
cursor.execute(query)<br>

#获取查询结果<br>
results = cursor.fetchall()<br>

#关闭游标和连接<br>
cursor.close()<br>
connection.close()<br>

#创建新的Word文档<br>
doc = Document()<br>

#设置标题<br>
doc.add_heading('老师上课时间表', level=1)<br>

#遍历查询结果并将数据写入表格<br>
table = doc.add_table(rows=1, cols=3)<br>
table.style = 'Table Grid'  # 设置表格样式<br>

#写入表头<br>
header_cells = table.rows[0].cells<br>
header_cells[0].text = '老师姓名'<br>
header_cells[1].text = '课程名称'<br>
header_cells[2].text = '时间'<br>

#写入查询结果<br>
for result in results:<br>
    row_cells = table.add_row().cells<br>
    row_cells[0].text = result[0]<br>
    row_cells[1].text = result[1]<br>
    row_cells[2].text = result[2]<br>

#保存Word文档<br>
doc.save('timetable.docx')<br>
