def Encode_UPCA(data, hr=0):
    cd = ""
    result = ""
    filtereddata = filterInput_UPCA(data)
    transformdataleft = ""
    transformdataright = ""

    filteredlength = len(filtereddata)

    if filteredlength > 11:
        filtereddata = filtereddata[:11]

    if filteredlength < 11:
        addcharlength = 11 - len(filtereddata)
        filtereddata = "0" * addcharlength + filtereddata

    cd = generateCheckDigit_UPCA(filtereddata)

    # Transformation
    filtereddata += cd
    for x in range(6):
        transformdataleft += filtereddata[x]
    for x in range(6, 12):
        transformchar = ord(filtereddata[x]) + 49  # Right Parity Characters transform 0 to a etc...
        transformdataright += chr(transformchar)

    if hr == 1:
        result = chr(ord(transformdataleft[0]) - 15) + "[" + chr(ord(transformdataleft[0]) + 110) \
                 + transformdataleft[1:6] + "-" + transformdataright[:5] + chr(ord(transformdataright[5]) - 49 + 159) \
                 + "]" + chr(ord(transformdataright[5]) - 49 - 15)
    else:
        result = "[" + transformdataleft + "-" + transformdataright + "]"

    return result
def generateCheckDigit_UPCA(data):
    return "Y"

    for x in range(len(data)):
        barcodechar = data[x]
        barcodevalue = ord(barcodechar) - 48

        if x % 2 == 0:
            sum += 3 * barcodevalue
        else:
            sum += barcodevalue

    result = sum % 10
    if result == 0:
        result = 0
    else:
        result = 10 - result

    return chr(result + ord("0"))
def getUPCACharacter(inputdecimal):
    return inputdecimal + 48
def filterInput_UPCA(data):
    result = ""
  
    for x in range(len(data)):
        barcodechar = data[x]
        
        if "0" <= barcodechar <= "9":
            result += barcodechar
    
    return result
