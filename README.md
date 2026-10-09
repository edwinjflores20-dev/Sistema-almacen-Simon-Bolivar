\# Sistema de Gestión de Almacén Hospitalario



\## Descripción



El \*\*Sistema de Gestión de Almacén Hospitalario\*\* es una aplicación web desarrollada en \*\*Python 3.x\*\* utilizando el framework \*\*Flask\*\*. Su finalidad es facilitar la gestión y control de los productos almacenados, permitiendo registrar productos, consultar el stock y gestionar los movimientos de inventario.



El sistema utiliza un archivo de Excel como medio de almacenamiento de la información, permitiendo conservar los datos de productos y movimientos realizados.



\## Objetivo



Desarrollar un sistema que permita mejorar el control de los productos de un almacén hospitalario mediante una aplicación web sencilla, organizada y fácil de utilizar.



\## Tecnologías utilizadas



\* Python 3.x

\* Flask

\* OpenPyXL

\* HTML

\* CSS

\* Microsoft Excel

\* GitHub



\## Funcionalidades principales



El sistema permite:



\* Registrar productos.

\* Consultar productos registrados.

\* Consultar el stock disponible.

\* Registrar movimientos de inventario.

\* Controlar las cantidades ingresadas.

\* Validar los datos registrados.

\* Detectar situaciones relacionadas con el stock mínimo.

\* Almacenar la información en un archivo Excel.

\* Acceder al sistema mediante un navegador web.



\## Estructura del proyecto



```text

sistema-almacen-hospital/

│

├── app.py

│

├── almacen\_hospital.xlsx

│

└── README.md

```



\### `app.py`



Contiene el código principal de la aplicación desarrollada en Python y Flask. En este archivo se encuentran las rutas, funciones, validaciones y operaciones necesarias para el funcionamiento del sistema.



\### `almacen\_hospital.xlsx`



Archivo utilizado para almacenar la información del sistema.



Contiene información relacionada con:



\* Productos.

\* Stock.

\* Stock mínimo.

\* Movimientos del almacén.



\### `README.md`



Documento que contiene la descripción, características, tecnologías y pasos necesarios para ejecutar el proyecto.



\## Instalación



Para ejecutar el proyecto se necesita tener instalado \*\*Python 3.x\*\*.



\### 1. Descargar o clonar el repositorio



Se puede descargar el proyecto desde GitHub o clonarlo mediante:



```bash

git clone URL\_DEL\_REPOSITORIO

```



Reemplazar `URL\_DEL\_REPOSITORIO` por la dirección correspondiente al repositorio.



\### 2. Ingresar a la carpeta del proyecto



```bash

cd sistema-almacen-hospital

```



\### 3. Instalar las dependencias



Ejecutar:



```bash

pip install flask openpyxl

```



\## Ejecución del sistema



Para iniciar la aplicación, ejecutar:



```bash

python app.py

```



Después de iniciar el servidor, Flask mostrará la dirección local donde se encuentra disponible la aplicación.



Normalmente se puede acceder desde:



```text

http://127.0.0.1:5000

```



También puede utilizarse:



```text

http://localhost:5000

```



\## Almacenamiento de información



El sistema utiliza el archivo:



```text

almacen\_hospital.xlsx

```



como medio de almacenamiento de los datos.



La biblioteca \*\*OpenPyXL\*\* permite que Python pueda leer y modificar la información contenida en el archivo Excel.



\## Validaciones y manejo de errores



El sistema incorpora validaciones para evitar el registro de información incorrecta.



Entre las principales validaciones se consideran:



\* Control de cantidades.

\* Evitar valores de stock negativos.

\* Validación de datos ingresados.

\* Control de movimientos de inventario.

\* Manejo de errores mediante excepciones.



Esto permite mejorar la integridad y confiabilidad de la información registrada.



\## Paradigmas de programación



El proyecto integra diferentes conceptos de programación.



\### Programación orientada a objetos



Se utilizan clases y métodos para representar elementos relacionados con el sistema, aplicando el concepto de encapsulamiento mediante atributos y métodos de acceso.



\### Programación funcional



Se emplean funciones de orden superior como:



```python

map()

filter()

reduce()

```



Estas permiten procesar colecciones de datos de manera más eficiente.



\### Manejo de excepciones



Se utilizan estructuras como:



```python

try:

&#x20;   ...

except:

&#x20;   ...

```



para controlar errores durante la ejecución del programa.



\## Base de datos utilizada



Para esta versión del proyecto se utiliza un archivo \*\*Excel\*\* como mecanismo de almacenamiento de información.



Archivo:



```text

almacen\_hospital.xlsx

```



Este archivo permite conservar los datos utilizados por la aplicación.



\## Requisitos



Para ejecutar correctamente el proyecto se requiere:



\* Windows, Linux o macOS.

\* Python 3.x.

\* Flask.

\* OpenPyXL.

\* Navegador web.

\* Archivo `almacen\_hospital.xlsx`.



\## Autor



\*\*Proyecto académico – Sistema de Gestión de Almacén Hospitalario\*\*



\## Repositorio



El código fuente del proyecto se encuentra publicado en GitHub para facilitar su revisión, almacenamiento y control de versiones.


## Instalación y ejecución

1. Instalar Python 3.
2. Instalar las dependencias con el comando:

   pip install flask openpyxl

3. Ejecutar el sistema:

   python app.py

4. Abrir el navegador en:

   http://127.0.0.1:5000

El sistema utiliza Flask para la interfaz web y OpenPyXL para gestionar los datos almacenados en Excel.

## Módulos del sistema

- Inicio: muestra un resumen del inventario.
- Productos: permite consultar los productos registrados.
- Nuevo producto: permite registrar productos.
- Movimientos: registra entradas y salidas.
- Historial: muestra los movimientos registrados.
- Alertas: identifica productos con stock crítico.

## Almacenamiento de datos

El sistema utiliza el archivo almacen_hospital.xlsx para guardar la información.

La hoja Productos contiene:
- ID del producto.
- Nombre.
- Categoría.
- Stock actual.
- Stock mínimo.

La hoja Movimientos contiene:
- ID del movimiento.
- ID del producto.
- Tipo de movimiento.
- Cantidad.
- Fecha.
- Motivo.

Se recomienda realizar copias de seguridad periódicas.

