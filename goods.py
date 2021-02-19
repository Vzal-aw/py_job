'''
shopping chart
'''
import re
goods = []
shopp_shar = []
#更改文件中内容的格式，并重新保存
with open('trade_name',encoding="utf-8") as f:
    for lin in f:
        lin = lin.split(None)
        dic = {"name":lin[0],"price":lin[1]}
        goods.append(dic)
print(goods)

#判断用户输入，并解决bug

for i in range(3):
    user_account = input("请输入金额：")
    ret = re.findall('[\D\s]',user_account)
    ret2 = ''.join(ret)
    if user_account.isdecimal():
        user_account = int(user_account)
        break
    elif user_account == ret2:
        print("请重新输入数字金额")

#打印商品及价格
count = []
for num, dic in enumerate(goods,1):
    count.append(num)
    print(f'{num} 名称：{dic["name"]} 价格：{dic["price"]}')

while 1:
    user_in = input("请选择商品序号：")
    if user_in.isdecimal():
        user_in = int(user_in)
    else:
        print('请输入数字序号')

    if user_in in count:
        print(f'您选择的是：{goods[user_in-1]}')
        shopp_shar.append(goods[user_in-1])
    elif user_in not in count:
        print('输入数值不在范围内，请重新输入')





