def RmbBigMaker(number):
    CnBigList = ['万', '仟', '佰', '拾', '']
    CnBigNum = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
    if not number.isdigit():
        return '请输入数字！'
    elif len(number)>5:
        return '请输入五位及以下的数字。'
    elif number[0] is '0':
        return '请输入正确数字！'
    flag =  5 - len(number)
    rmb = ''
    zero_flag = True
    for i in number:
        if i=='0' and flag<5:
            zero_flag = False
        else:
            if not zero_flag:
                rmb = rmb + CnBigNum[int(0)]
            rmb = rmb + CnBigNum[int(i)]+CnBigList[flag]
            zero_flag = True
        flag +=1
        
    return rmb
    
