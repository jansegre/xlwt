import six

if six.PY3:
    long = int

    def byteindex(data, index):
        return ord(data[index])

    def iterbytes(data):
        return (chr(char).encode('latin-1') for char in data)
else:
    long = long
    byteindex = lambda x, i: x[i]
    iterbytes = lambda x: iter(x)
