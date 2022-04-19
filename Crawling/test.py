def foo(n):
    s = [n]  # set as the list
    def bar(i):
        s[0] += i
        return s[0]
    return bar(1)



if __name__ == '__main__':
    a = foo(2)
    # foo1(2)