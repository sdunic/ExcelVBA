def luhn_algorithm(luhnInput):
    luhnSum = 0
    for i in range(len(luhnInput)-2, -1, -1): 
        if i % 2 == len(luhnInput) % 2:
            if 2 * int(luhnInput[i]) > 9:
                luhnSum += (2* int(luhnInput[i]) % 10)
                luhnSum += (2* int(luhnInput[i]) // 10)
            else:
                luhnSum += 2 * int(luhnInput[i])
        else:
            luhnSum += int(luhnInput[i])

    checkDigit = (luhnSum*9) % 10

    if checkDigit == int(luhnInput[-1]):
        return "VALID"
    else:
        return checkDigit
    return luhnSum

def single_digit_error(luhnInput):
    luhnOutputUp = ""
    luhnOutputDown = ""
    
    for i in range(len(luhnInput)-1):
        for k in range(len(luhnInput)):
            if i == k:
                luhnOutputUp += str(int(luhnInput[k]) + 1)
                if(int(luhnInput[k]) - 1 >= 0):
                    luhnOutputDown += str(int(luhnInput[k]) - 1)
                else:
                    luhnOutputDown += str(int(luhnInput[k]))
            else:
                luhnOutputUp += luhnInput[k]
                luhnOutputDown += luhnInput[k]

        if(luhn_algorithm(luhnOutputUp) == "VALID"):
            return luhnOutputUp
        elif (luhn_algorithm(luhnOutputDown) == "VALID"):
            return luhnOutputDown
        else:
            luhnOutputUp = ""
            luhnOutputDown = ""

def transposition_error(luhnInput):
    luhnOutput = ""
    for i in range(len(luhnInput)-1):
        for k in range(len(luhnInput)):
            if(k == i):
                   luhnOutput += (luhnInput[i+1] + luhnInput[i])
            elif(k == i+1):
                   continue
            else:
                luhnOutput += luhnInput[k]
        if luhn_algorithm(luhnOutput) == "VALID":
            return luhnOutput
        else:
            luhnOutput = ""

def twin_error(luhnInput):
    luhnOutputUp = ""
    luhnOutputDown = ""
    doubleDigitsFlag = False
    for i in range(len(luhnInput)-1):
        for k in range(len(luhnInput)):
            if(i == k and luhnInput[i] == luhnInput[i+1]):
                doubleDigitsFlag = True
                luhnOutputUp += str(int(luhnInput[k]) + 1) + str(int(luhnInput[k]) + 1)
                if(int(luhnInput[k]) - 1 >= 0):
                    luhnOutputDown += str(int(luhnInput[k]) - 1) + str(int(luhnInput[k]) - 1)
                else:
                    luhnOutputDown += str(int(luhnInput[k])) + str(int(luhnInput[k]))
            elif (k == i + 1):
                continue
            else:
                luhnOutputUp += luhnInput[k]
                luhnOutputDown += luhnInput[k]

        if(doubleDigitsFlag and luhn_algorithm(luhnOutputUp) == "VALID"):
            return luhnOutputUp
        elif (doubleDigitsFlag and luhn_algorithm(luhnOutputDown) == "VALID"):
            return luhnOutputDown
        else:
            luhnOutputUp = ""
            luhnOutputDown = ""
            doubleDigitsFlag = False



for i in range(4):
    print(luhn_algorithm(input()))
for i in range(2):
    print(single_digit_error(input()))
for i in range(2):
    print(transposition_error(input()))
for i in range(2):
    print(twin_error(input()))






















