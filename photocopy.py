
class PhotocopyManager:            
    current_log_file = None
    total = 0
    current_user = "Marta"
    current_print_type = "black_and_white"
    initial_value_in_the_box = 0
    previous_added_price = 0
    def init_logger(self):
        print("Init")
        if not os.path.isdir("./datos"):
            os.mkdir("./datos")

        today = date.today()
        label_date = builder.get_object("label_date")
        label_date.set_text(str(today))
        filepath = "./datos/" + str(today) + ".txt" 
        if os.path.isfile(filepath):
            readed_file = open("./datos/"+str(today)+".txt","r")
            lines = readed_file.readlines()
    current_print_type = "black_and_white"
    initial_value_in_the_box = 0
    previous_added_price = 0
    def init_logger(self):
        print("Init")
        if not os.path.isdir("./datos"):
            os.mkdir("./datos")

        today = date.today()
        label_date = builder.get_object("label_date")
        label_date.set_text(str(today))
        filepath = "./datos/" + str(today) + ".txt" 
        if os.path.isfile(filepath):
            readed_file = open("./datos/"+str(today)+".txt","r")
            lines = readed_file.readlines()
            position = lines[0].find("TOTAL: ")
            #print(lines[0][position+7])
            offset1=position+7
            offset2=len(lines[0])-4
            result=lines[0][offset1:offset2]
            intresult=result.replace(",", "")
            new_string = ''.join(e for e in intresult if e.isalnum())
            re.sub('[^A-Za-z0-9]+', '', new_string)
            self.total=int(new_string)
        else:
            new_file = open("./datos/"+str(today)+".txt","w+")
            new_file.write("Fecha: ")
            new_file.write(str(today) + "                                ")
            new_file.write("TOTAL: 0 Gs\n")
            new_file.write("Hora                 Cantidad                Tipo\n")
            new_file.close()
