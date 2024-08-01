from util.archivoPrimero import ArchivoPrimero
#from util.archivoSegundo import ArchivoSegundo
from util.archivoTercero import ArchivoTercero
from util.archivoCuarto import ArchivoCuarto
#from util.archivoSexto import ArchivoSexto
from util.archivoQuinto import ArchivoQuinto
from util.archivoSeptimo import ArchivoSeptimo
from util.archivoOctavo import ArchivoOctavo
from util.archivoNoveno import ArchivoNoveno

def main():

    ArchivoPrimero().crearArchivoPrimero()
    #ArchivoSegundo().crearArchivoSegundo()
    ArchivoTercero().crearArchivoTercero()
    ArchivoCuarto().crearArchivoCuarto()
    ArchivoQuinto().crearArchivoQuinto()
    #ArchioSexto().crearArchivoSexto()
    ArchivoSeptimo().crearArchivoSeptimo()
    ArchivoOctavo().crearArchivoOctavo()
    ArchivoNoveno().crearArchivoNoveno()



if __name__ == "__main__":
    main()