def ordinal(n):
    # Converte um numeral em texto ordinal através de um dicionário, a princípio.
    ord = {
        1: "Primeiro",
        2: "Segundo",
        3: "Terceiro"
    }
    return ord.get(n)