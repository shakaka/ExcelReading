
def findLen(str):
    counter = 0
    while str[counter:]:
        counter += 1
    return counter

str = "GG有中文152G來有"
if len(str) > 5:
    str = str[:5]
print(str)
print(findLen(str))
print(len(str))
