import sys


def reverseWords(input):
    # 通过空格将字符串分隔符，把各个单词分隔为列表
    inputWords = input.split(" ")

    # 翻转字符串
    # 假设列表 list = [1,2,3,4],
    # list[0]=1, list[1]=2 ，而 -1 表示最后一个元素 list[-1]=4 ( 与 list[3]=4 一样)
    # inputWords[-1::-1] 有三个参数
    # 第一个参数 -1 表示最后一个元素
    # 第二个参数为空，表示移动到列表末尾
    # 第三个参数为步长，-1 表示逆向
    inputWords = inputWords[-1::-1]

    print(isinstance(inputWords,list))

    # 重新组合字符串
    output = ' '.join(inputWords)

    return output

def FibonacciSeries():
    a, b = 0, 1
    while b < 10:
        print(a,b)
        a, b = b, a+b

def iterTest():
   arr = [3,4,8,7,9,6,6]
   it = iter(arr)
   for x in it:
       print(x,end=" ")

# 不定长参数
def printinfo(arg1, *vartuple):
    "打印任何传入的参数"
    print("输出: ")
    print(arg1)
    print(vartuple)

# 可写函数说明
def printinfo2(arg1, **vardict):
    "打印任何传入的参数"
    print("输出: ")
    print(arg1)
    print(vardict)

if __name__ == "__main__":
    # input = '1 2 3 4 5 6 7 8 9'
    # rw = reverseWords(input)

    # print(rw)
    # FibonacciSeries()
    # iterTest()

    # 调用printinfo 函数
    printinfo(70, 60, 50)

    # 调用printinfo 函数
    printinfo2(1, a=2, b=3)

    for i in sys.argv:
        print(i)


