num1 = int(input("número 1: "))
num2 = int(input("número 2: "))
operacion = input("operación (+, -, *, /): ")

match operacion:
    case "+":
        res = num1 + num2
    case "-":
        res = num1 - num2
    case "*":
        res = num1 * num2
    case "/":
        res = num1 / num2
    case _:
        print("Operación no válida.")
        
print(f"el resultado de {num1} {operacion} {num2} es {res}")