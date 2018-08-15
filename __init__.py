# -*- coding: utf-8 -*-

dictionary = {0: "", 1: "", 2: "тысяч", 3: "миллион", 4: "миллиард", 5: "триллион"}

dicthundreds = {"1": "сто", "2": "двести", "3": "триста",
    "4": "четыреста",
    "5": "пятьсот",
    "6": "шестьсот", "7": "семьсот", "8": "восемьсот",
    "9": "девятьсот"}

dictdecimals = {"2": "двадцать", "3": "тридцать",
     "4": "сорок", "5": "пятьдесят", "6": "шестьдесят", "7": "семьдесят",
     "8": "восемьдесят",
     "9": "девяносто"}

dictunits = {"1": "один", "2": "два", "3": "три", "4": "четыре", "5": "пять",
    "6": "шесть", "7": "семь",
    "8": "восемь", "9": "девять", "01": "один", "02": "два", "03": "три",
    "04": "четыре", "05": "пять", "06": "шесть", "07": "семь",
    "08": "восемь", "09": "девять", "10": "десять", "11": "одиннадцать",
    "12": "двенадцать",
    "13": "тринадцать", "14": "четырнадцать", "15": "пятьнадцать",
    "16": "шестьнадцать",
    "17": "семнадцать", "18": "восемнадцать", "19": "девятнадцать"}


def subnumb(input):
    result = ""
    if input == "000":
        return result
    if input[0] != "0":
        result += " " + dicthundreds[input[0]]
    if (input[1] != "0" and input[1] != "1"):
        result += " " + dictdecimals[input[1]]
        result += " " + dictunits[input[2]]
        return result
    suffix = input[1] + input[2]
    result += " " + dictunits[suffix]
    return result


def numbers(input_):
    input  = prepared(input_)
    groups = len(input) // 3
    print("groups", groups)

    print(dictionary[groups])
    result = ""

    iter=0
    while (groups!=0):

        prefix = input[iter:iter+3]
        print("prefix=", prefix)
        iter += 3
        if subnumb(prefix) == "":
            continue
        print("grops", groups)
        print("dictionary[groups]", dictionary[groups])
        result += subnumb(prefix) + dictionary[groups]

        groups -= 1

    print(result)

    return result


def prepared(input):
    rest= 3 - (len(input) % 3)
    if (rest == 3):
        return input
    result = ""

    while (rest != 0):
        rest -= 1
        result += "0"

    result += input

    print("result", result)
    return result

numbers("3454334")