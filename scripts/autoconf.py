import os
import pip
try:
    import xlrd
except ImportError:
    pip.main(['install', "xlrd"])   

wb = xlrd.open_workbook("wb.xls")
sheet = wb.sheet_by_index(0)

help = """<<<help>>>
help - pokazuje dostepne opcje
define - definiowanie hostów
services - definiowanie usług
ping - definiowanie usługi ping
exit - opuszcza program
"""
hosts = """<<<hosts>>>
help - pokazuje dostepne rodzaje hostów
windows - windows-server
camera - generic-camera
printer - generic-printer
back - opuszcza sekcję definicji hostów
exit - opuszcza program
"""
services = """<<<services>>>
help - pokazuje dostepne rodzaje usług
server - usługi serwerowe
printer - usługi drukarek
back - opuszcza sekcję definicji usług
exit - opuszcza program
"""
server_services = """<<<server services>>>
help - pokazuje dostepne rodzaje hostów
windows - windows-server
camera - generic-camera
printer - generic-printer
back - opuszcza sekcję definicji hostów
exit - opuszcza program
"""
printer_services = """<<<printer services>>>
help - pokazuje dostepne rodzaje hostów
windows - windows-server
camera - generic-camera
printer - generic-printer
back - opuszcza sekcję definicji hostów
exit - opuszcza program
"""

os.system("CLS")
print("""<<<UWAGA>>>
Program usunie wszystkie utworzone wcześniej pliki, sprawdź czy nie są już potrzebne.
Naciśnij ENTER żeby zatwierdzić.
"""
)
check = input()
files = ["camera.txt", "windows.txt", "printer.txt", "ping.txt", "server_services.txt"]

if check in "":
    for file in files:
        try:
            os.remove(file)
        except:
            pass
    os.system("CLS")
else:
    os.exit()

option = None
passed = ["exit", "back", "help"]
while not (option == "exit"):
    print(help)
    print("Podaj nazwe templatu: ")
    option = input()
    os.system("CLS")

    if option == "define":
        options = ["windows","camera","printer"]
        while not (option == "exit" or option == "back"):
            print(hosts)
            print("Podaj rodzaj hostów do zdefiniowania: ")
            option = input()
            os.system("CLS")

            if option == "windows":
                use = "windows-server"
            elif option == "camera":
                use = "generic-camera"
            elif option == "printer":
                use = "generic-printer"  
            else: 
                if option not in passed:
                    print('\n"'+option+'" - taki rodzaj hostów nie istnieje\n')     
                     
            if option in options:
                with open(option+".txt", "w") as list:
                    for pos in range(sheet.nrows):
                        ipaddress = sheet.cell_value(pos,0)
                        hostname = sheet.cell_value(pos,1)

                        list.write("define host{\n")
                        list.write("\tuse\t\t"+use+"\n")
                        list.write("\thost_name\t" + hostname + "\n")
                        list.write("\talias\t\t"+hostname+"\n")
                        list.write("\taddress\t\t"+ipaddress+"\n")
                        if option == "printer":
                            list.write("\thostgroups\tnetwork-printers\n")
                            list.write("\tparents\t\tTYC-PRINT-2\n")
                        if option == "camera":
                            list.write("\thostgroups\tIPCAM\n")
                        list.write("\t}\n")
        
                list.close()
                print("\nkonfiguracja zapisana do "+option+".txt\n")

    elif option == "services":
        options = ["server","printer"]
        while not (option == "exit" or option == "back"):
            print(services)
            print("Podaj rodzaj usług do zdefiniowania: ")
            option = input()
            os.system("CLS")

            if option == "server":
                if option in options:
                    with open(option+".txt", "w") as list:
                        for pos in range(sheet.nrows):
                            ipaddress = sheet.cell_value(pos,0)
                            hostname = sheet.cell_value(pos,1)

                            list.write("define host{\n")
                            list.write("\tuse\t\t"+use+"\n")
                            list.write("\thost_name\t" + hostname + "\n")
                            list.write("\talias\t\t"+hostname+"\n")
                            list.write("\taddress\t\t"+ipaddress+"\n")
                            if option == "printer":
                                list.write("\thostgroups\tnetwork-printers\n")
                                list.write("\tparents\t\tTYC-PRINT-2\n")
                            if option == "camera":
                                list.write("\thostgroups\tIPCAM\n")
                            list.write("\t}\n")

                    list.close()
                    print("\nkonfiguracja zapisana do "+option+".txt\n")

            elif option == "printer":
                if option in options:
                    with open(option+".txt", "w") as list:
                        for pos in range(sheet.nrows):
                            ipaddress = sheet.cell_value(pos,0)
                            hostname = sheet.cell_value(pos,1)
    
                            list.write("define host{\n")
                            list.write("\tuse\t\t"+use+"\n")
                            list.write("\thost_name\t" + hostname + "\n")
                            list.write("\talias\t\t"+hostname+"\n")
                            list.write("\taddress\t\t"+ipaddress+"\n")
                            if option == "printer":
                                list.write("\thostgroups\tnetwork-printers\n")
                                list.write("\tparents\t\tTYC-PRINT-2\n")
                            if option == "camera":
                                list.write("\thostgroups\tIPCAM\n")
                            list.write("\t}\n")
            
                    list.close()
                    print("\nkonfiguracja zapisana do "+option+".txt\n")
                    
            else: 
                if option not in passed:
                    print('\n"'+option+'" - taki rodzaj usług nie istnieje\n')     
                     
            

    elif option == "ping":
        with open(option+".txt", "w") as list:
            for pos in range(sheet.nrows):
                hostname = sheet.cell_value(pos,1)

                list.write("define service{\n")
                list.write("\tuse\t\t\tgeneric-service\n")
                list.write("\thost_name\t\t" + hostname + "\n")
                list.write("\tservice_description\tPING\n")
                list.write("\tcheck_command\t\tcheck_ping!200.0,20%!400.0,60%\n")
                list.write("\tcheck_interval\t\t5\n")
                list.write("\tretry_interval\t\t1\n")
                list.write("\t}\n")

            list.close()
            print("\nkonfiguracja zapisana do "+option+".txt\n")

    else: 
        if option not in passed:
            print('\n"'+option+'" - taki template nie istnieje\n')
            