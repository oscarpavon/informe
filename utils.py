
def formated_namber(value):
    formated = "{:,}".format(value)
    formated = formated.replace(',','.')
    return formated 

def today_file(read_format):
    today = date.today()
    filepath = "./datos/" + str(today) + ".txt"
    from_file = open(filepath,read_format)

def get_text_moth(number):
    switcher = {
            1: "Enero",
            2: "Febrero",
            3: "Marzo",
            4: "Abril",
            5: "Mayo",
            6: "Junio",
            7: "Julio",
            8: "Agosto",
            9: "Setiembre",
            10: "Octubre",
            11: "Noviembre",
            12: "Diciembre"
            }
    return switcher.get(number,"Invalid")
