def columnToLetter(column):
    temp, letter = '', ''
    while column > 0:
        temp = (column - 1) % 26
        letter = chr(temp + 65) + letter
        column = (column - temp - 1) // 26
    return letter


def letterToColumn(letter):
    column = 0
    for i in range(len(letter)):
        column += (ord(letter[i]) - 64) * pow(26, len(letter) - i - 1)
    return column
