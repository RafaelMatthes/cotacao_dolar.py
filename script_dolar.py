from classes import Dolar, Arquivo

dolar = Dolar()
year,month,last_day = dolar.get_range_datas()
lista = dolar.get_dict_data(year,month,last_day)

arquivo = Arquivo()
arquivo.build_xlsx(lista)
arquivo.set_valores_log(arquivo.load_from_xlsx())
arquivo.salva_txt()
