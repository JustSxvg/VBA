# VBA
Excel VBA Macros
Teniendo los precios a los que diariamente cotiza una acción, la siguiente macro analiza mediante ciertas condiciones si "Comprar" o "Vender", y tras tomar esta decisión, si "Comprar pronto" o "Vender pronto". Para esto, hay un plazo de 30 días en los que se visualiza como se comporta la cotización y, a partir del día 30, se empiezan a tomar estas decisiones de compra o venta.

Tenemos 4 valores que son útiles para este análisis:
- p10 : Promedio de los últimos 10 días del precio de cierre diario para dicha acción
- p30 : Promedio de los últimos 30 días del precio de cierre diario para dicha acción
- closet : Valor del primer día desde el que se tomará una decisión
- closetmenos1: Valor del anterior anterior al que se tomará la primera decisión

También tenemos las siguientes condiciones para tomar una decisión:
1. Si p10 > p30, entonces "Comprar"
2. Si p10 < p30, entonces "Comprar en corto"
3. Si closet > closetmenos1*0.8, entonces "Vender"
4. Si closet < closetmenos1*0.8, entonces "Vender en corto"

UNMSM, FCE
