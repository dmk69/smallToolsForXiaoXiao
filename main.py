# 此程序需要安装python及xlwings
# 最后修改 2021/07/24
# 作者 王德军
# 著作权归属 赵潇颂
import xlwings as xw

# 将下列数值更改为每页中你需要的行数
youNeedRow = input('请输入你需要的每页行数,并按回车键：')

# 将你的Excel表格真实地址复制到下面的单引号中
filenames = input('请将你的Excel文件绝对位置复制至此,并按回车键：')
wb = xw.Book(filenames)

# 将单引号中Sheet0更改为你表格中的表单名称
sht = wb.sheets['Sheet0']

print(wb.fullname)
# 将数据转义至新的表格

headData = sht.range('A1:D1').value
print(headData)

cell = sht.used_range.last_cell
rows = cell.row
colums = cell.column
print('总行数 '+str(rows)+'总列数 '+str(colums))

pagesNeeds=int((rows-1)/(int(youNeedRow)-1))+1
print('需要页码: '+ str(pagesNeeds))

acrow = 2
for i in range(1,pagesNeeds+1):
    nb = xw.Book()
    sht1 = nb.sheets[0]

    sht1.range('A1').value = headData
    sht.range((acrow,1),(acrow+(int(youNeedRow)-2),4)).api.Copy(sht1.range('A2').api)
    acrow = acrow+(int(youNeedRow)-1)

    nb.save()
    nb.close()

# nb = xw.Book()
# sht1 = nb.sheets[0]

# sht.range('E48').api.Copy(sht1.range('A2').api)
# sht1.range('A1').value = headData
# sht.range('A2:D2000').api.Copy(sht1.range('A2').api)
# A2 D2000 A2001 D3999 A4000 D5998 A5999 D7997
# A2 A(2+1999)=A2001 A(2001+1999)=
# A 增加1999 D增加1998