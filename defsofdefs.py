def Zero_adder(n):
    try:
        if int(n) > 9:
            return f'{n}'
        else:
            return f'0{n}'
    except:
        print('Ocorreu um erro na função Zero_adder em defsofdefs.py')