# results = [x if x%2==0 else 'TEK' for x in range(1,10)]
# print(results)

# list = [1,2,3,5]

# list.append(32)

# print(list)

# def total(number1,number2):
#     return number1*number2

# print(total(5,10))


# def numbers(num1):
#     x=[]

#     for i in range(1,num1):
#         if num1%i == 0:
#             x.append(i)
#     return x

# print(numbers(20))

def square(number):
    return number**2
number = [3,6,9,10]
result = list(map(square,number))

print(result)


