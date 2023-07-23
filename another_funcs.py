def be_nums(num):
    ans = ''
    dr = round(num - int(num), 1)
    num = int(num)
    for i in range((len(str(num)) - 1) // 3 + 1):
        ln = str(num % (1000 ** (i + 1)) // (1000 ** i))
        ln = '0' * (3 - len(ln)) + ln
        ans = ln + " " + ans
    ans = ans.rstrip().lstrip('0')
    if ans == '': ans = '0'
    return ans

# print(be_nums(1000))
