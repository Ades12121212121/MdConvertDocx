# MdConvertDocx

MdConvertDocx es una herramienta avanzada para convertir archivos Markdown en documentos DOCX con formato enriquecido. Utiliza diversas bibliotecas de Python para ofrecer una experiencia de conversión completa y personalizable.

## Características

- **Soporte para Markdown completo**: Reconoce encabezados, listas, citas, enlaces, imágenes, texto en negrita y cursiva, código en línea y más.
- **Temas personalizables**: Permite seleccionar diferentes temas para el documento DOCX de salida (por ejemplo, "default" y "professional").
- **Interfaz interactiva**: Utiliza la biblioteca Rich para proporcionar una interfaz de usuario amigable y atractiva en la línea de comandos.
- **Corrección de codificación**: Soluciona problemas comunes de codificación en los archivos Markdown.

## Requisitos

- Python 3.6 o superior
- Dependencias de Python (especificadas en `requirements.txt`)

## Instalación

1. Clona este repositorio:

    ```sh
    git clone https://github.com/Ades12121212121/MdConvertDocx.git
    cd MdConvertDocx
    ```

2. Instala las dependencias:

    ```sh
    pip install -r requirements.txt
    ```

## Uso

Puedes ejecutar la herramienta directamente desde la línea de comandos. Aquí hay algunos ejemplos de uso:

### Conversión básica

Para convertir un archivo Markdown a DOCX con el tema por defecto:

```sh
python main.py -i ruta/a/tu/archivo.md -o salida.docx
