# -*- coding: utf-8 -*-

dictionary = {1: "тысяч", 2: "миллион", 3: "миллиард"}

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
    "17": "семьнадцать", "18": "восемнадцать", "19": "девятнадцать"}


def subnumb(input):
    result = ""
    if input[0] != "0":
        result += dicthundreds[input[0]]
    if (input[1] != "0" and input[1] != "1"):
        result += dictdecimals[input[1]]
        result += dictunits[input[2]]
        return result
    suffix = input[1] + input[2]
    result += dictunits[suffix]
    return result


def numbers(input):
    digest = len(input)
    counter = len(input) // 3
    result = ""
    result += dictionary[counter]
    return result

print(subnumb("123"))
print(subnumb("466"))
print(subnumb("103"))
print(subnumb("999"))
print(numbers("1000000000"))
print(numbers("1000000"))
print(numbers("1156"))
assert numbers("156") == "сто пятьдесят шесть"