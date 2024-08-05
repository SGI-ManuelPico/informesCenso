from util.archivoPrimero import ArchivoPrimero
from util.archivoSegundo import ArchivoSegundo
from util.ArchivoTercero import ArchivoTercero
from util.ArchivoCuarto import ArchivoCuarto
from util.archivoSexto import ArchivoSexto
from util.archivoQuinto import ArchivoQuinto
from util.ArchivoSeptimo import ArchivoSeptimo
from util.ArchivoOctavo import ArchivoOctavo
from util.archivoNoveno import ArchivoNoveno

def main():

    ArchivoPrimero().crearArchivoPrimero()
    ArchivoSegundo().crearArchivoSegundo()
    ArchivoTercero().crearArchivoTercero()
    ArchivoCuarto().crearArchivoCuarto()
    ArchivoQuinto().crearArchivoQuinto()
    ArchivoSexto().crearArchivoSexto()
    ArchivoSeptimo().crearArchivoSeptimo()
    ArchivoOctavo().crearArchivoOctavo()
    ArchivoNoveno().crearArchivoNoveno()



if __name__ == "__main__":
    main()