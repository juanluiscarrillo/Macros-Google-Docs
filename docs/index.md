## Macros para Google Docs

En este repositorio se colgarán macros realizadas para Google Doc.

En este momento se encuentran disponibles las siguientes macros:
- Valores

En los siguientes apartados se puede obtener más información sobre las macros

## Valores

Se trata de un conjunto de macros para hojas de cálculo y formularios que gestiona las inversiones de un usuario. A continuación, se detallan los documentos presentes en el proyecto:
- Carpeta ***Empresas***: En esta carpeta habrá tantos documentos de hoja de cálculo como valores en los que se ha invertido o se pretende invertir. Además, hay una hoja de cálculo con nombre ***00 AA Plantilla*** que se utilizará como plantilla para crear cada hoja excel de cada valor.
- Hoja de cáluculo ***Resumen***: La hoja principal llamada ***Inversión*** contendrá un resumen con todos los valores en los que se ha invertido. Además, puede haber más hojas cuyo nombre coincide con el número correspondente al año, donde se desglosan las operaciones realizadas en el ejercicio correspondiente.
- Formulario ***Valores***: Es el formulario que se empleará para introducir las operaciones de compra y venta, así como los dividendos.

Antes de utilizar este software es importante asegurarse previamente de:
- El hipervínculo de la celda **C1** del fichero Resumen se corresponde con la URL de la *Vista Previa" del formulario
- La variable **GLOBAL.formId** del script del formulario se corresponde con la ID del formulario
- La variable **GLOBAL.summaryId** del script del formulario se corresponde con la ID del documento *Resumen*

**NOTA:** La ID de un documento de Google Doc se puede obtener fácilmente de la URL. Por ejemplo, si se tiene una URL *https://docs.google.com/forms/d/1OkeqoNVFT4vPEl5sm-34vYeuPQyXpD1b0qetIFh8eBA/edit*, la ID es *1OkeqoNVFT4vPEl5sm-34vYeuPQyXpD1b0qetIFh8eBA*

Además, por cada nuevo valor se procederá de la siguiente manera:
- Dentro de la carpeta *Empresas* se crea una copia del documento ***00 AA Plantilla*** y se le cambia el nombre al documento con uno identificativo del valor.
- Se copia la URL del documento recien creado y se pega en la primera celda libre (es importante que no queden celdas vacías) de la columna ***A*** de la hoja *Inversión* del documento *Empresas*. 
- Se asigna un nombre al valor en la celda de la columna ***B*** situada justo a la derecha de la URL
- Se abre el formulario ***Valores*** en modo de edición y se añade un nuevo item en el campo desplegable **Valor**. El valor de dicho item debe coincidir exactamente con el nombre que se le dio al valor en el documento *Resumen* (celda de la columna B).

Una vez que se ha configurado el valor, se puede empezar a trabajar con él. Para ello, desde la hoja *Inversión* del documento *Resumen* se pulsa en el enlace de la celda *C1* para abrir el formulario de introducción de datos. Se completa el formulario y se envía. El programa automáticamente actualiza el contenido del documento de la carpeta *Empresas* correspondiente.

Para comprobar los cambios en la hoja *Inversión* del documento *Resumen*, se pulsa el botón *Actualizar* de la celda *A1*. El programa recoge todo los cambios previamente introducidos.

Como se ha comentado previamente, además de la hoja *Inversión* puede haber otras hojas en el documento *Resumen*. Cada una de estas hojas se corresponde con un ejercicio y en ella se mostraría un resumen correspondiente al mismo. Añadir más ejercicios es tan simple como duplicar la hoja de un año y cambiar el nombre de la hoja recién creada con el número correspondiente (no olvidar pulsar el botón *Actualizar* para recoger los cambios). Por último, si se quiere que se mostrar la información de ejercicio individual también en la hoja *Inversión*, hay que poner el número del año correspondiente en una celda de la fila 2 en la columna inmediatamente libre a la derecha, siempre, en cualquier caso, a la derecha de la letra k (no olvidar pulsar el botón *Actualizar* para recoger los cambios).
