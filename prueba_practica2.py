
#1
def contar_vocales(texto):
    vocales = 'aeiou'
    conteo = {vocal: 0 for vocal in vocales}
    for char in texto.lower():
        if char in conteo:
            conteo[char] += 1
    return conteo

print(contar_vocales("Hola Mundo"))

#2
def filtrar_pares(numeros):
    return [num for num in numeros if num % 2 == 0]

print(filtrar_pares([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]))

#3
def invertir_cadena(texto):
    return texto[::-1]

print(invertir_cadena("python"))

#4
def calcular_precio_final(precio, descuento):
    return precio * (1 - descuento / 100)

print(calcular_precio_final(100, 15))

#5
class Coche:
    def __init__(self, marca, modelo, año):
        self.marca = marca
        self.modelo = modelo
        self.año = año
        self.velocidad = 0

    def acelerar(self, cantidad):
        self.velocidad += cantidad

    def frenar(self, cantidad):
        self.velocidad = max(0, self.velocidad - cantidad)

    def mostrar_informacion(self):
        print(f"Marca: {self.marca}")
        print(f"Modelo: {self.modelo}")
        print(f"Año: {self.año}")
        print(f"Velocidad: {self.velocidad} km/h")

mi_coche = Coche("Toyota", "Corolla", 2020)
mi_coche.acelerar(20)
mi_coche.frenar(5)
mi_coche.mostrar_informacion()