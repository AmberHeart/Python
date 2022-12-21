while True:
    code = input("输入密码：")
    string = list(code)
    for x in range(26):
        tmp = ''
        for y in range(len(code)):
            if string[y] != ' ':
                temp = ord(string[y]) + x + 1
                if temp > 90 and ord(string[y]) <= 90:
                    temp -= 26
                if temp > 122:
                    temp -= 26
                tmp += '%c' % temp
            else:
                tmp += ' '
        print(tmp)
    for x in range(26):
        tmp = ''
        for y in range(len(code)):
            if string[y] != ' ':
                temp = ord(string[y]) - x - 1
                if temp < 65:
                    temp += 26
                if ord(string[y]) >= 97 and temp < 97:
                    temp += 26
                tmp += '%c' % temp
            else:
                tmp += ' '
        print(tmp)