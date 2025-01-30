import os
import re
from collections import Counter

# 1. Contar palabras únicas
def contar_palabras_unicas(texto):
    texto = re.sub(r'[^\w\s]', '', texto.lower())  
    palabras = texto.split()
    return dict(Counter(palabras))

entrada = "hola mundo, hola python"
resultado = contar_palabras_unicas(entrada)
print(resultado)

# 2. Verificar si un número es primo
def es_primo(numero):
    if numero <= 1:
        return False
    for i in range(2, int(numero ** 0.5) + 1):
        if numero % i == 0:
            return False
    return True

print(es_primo(7))
print(es_primo(10))

# 3. Calcular estadísticas de una lista de números
def calcular_estadisticas(numeros):
    if not numeros:
        return {"promedio": 0, "maximo": None, "minimo": None}
    promedio = round(sum(numeros) / len(numeros), 2)  
    return {"promedio": promedio, "maximo": max(numeros), "minimo": min(numeros)}

numeros = [4, 7, 1, 10, 9]
resultado = calcular_estadisticas(numeros)
print(resultado)

# 4. Procesar archivo (leer y escribir líneas en orden inverso)
def procesar_archivo(entrada, salida):
    try:
        if not os.path.isfile(entrada):
            raise FileNotFoundError(f"{entrada} no es un archivo válido.")
        
        with open(entrada, 'r', encoding='utf-8') as file:
            lineas = file.readlines()
        
        with open(salida, 'w', encoding='utf-8') as file:
            file.writelines(reversed(lineas))  
        
        print("Archivo procesado correctamente.")

    except FileNotFoundError:
        print(f"El archivo {entrada} no se encontró.")
    except PermissionError:
        print(f"No tienes permisos para acceder al archivo {entrada}.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")

entrada = r"U:\python\extraccion-sharepoint\entrada\entrada.txt"
salida = 'salida.txt'
procesar_archivo(entrada, salida)

# 5. Clase Empleado
class Empleado:
    def __init__(self, nombre, edad, salario):
        self.nombre = nombre
        self.edad = edad
        self.salario = salario
        
    def aumentar_salario(self, porcentaje):
        self.salario += self.salario * (porcentaje / 100)
     
    def mostrar_informacion(self):
        print(f"Nombre: {self.nombre}")
        print(f"Edad: {self.edad}")
        print(f"Salario: {self.salario:.2f}")

empleado = Empleado("Juan", 30, 50000)
empleado.mostrar_informacion()
empleado.aumentar_salario(10)
empleado.mostrar_informacion()

# 6. Inventario de productos
def productos():
    return [
        {"nombre": "Laptop", "precio": 800, "cantidad": 10},
        {"nombre": "Mouse", "precio": 20, "cantidad": 50},
        {"nombre": "Teclado", "precio": 30, "cantidad": 20}
    ]

def calcular_valor_total(productos):
    return sum(p["precio"] * p["cantidad"] for p in productos)  

inventario = productos()
print(f"Valor total del inventario: {calcular_valor_total(inventario)}")