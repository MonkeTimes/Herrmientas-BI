# Inicio
Repositorio que va a almacenar soluciones a utilizar en la empresa seleccionada para el proyecto de la materia taller de productividad basada en herramientas tecnologicas

# Resumen ejecutivo
### Descripcion
En este repositorio se encuentran los codigos para dos herramientas las cuales tienen como proposito apoyar en el proceso de limpieza de informacion del area de BI dentro de la empresa whirlpool, las herramientas que se realizaron son el codigo de google script que permite automatizar correos y una macro que permite juntar diferentes archivos en uno solo.

### Problema identificado
El problema que se identifico dentro del area de BI es que nos toma mucho tiempo procesar toda la informacion que nos llega de informacion de sell out e inventario de los diferentes clientes, algunos clientes nos hacen descargar de portales desde los cuales nos dan como resultado mas de 30 archivos los cuales tenemos que revisar de uno tomandonos horas adicionales, el otro problema es que tenemos uqe mandar correos de forma semanal a todos los KAMs que llevan las cuentas de los clientes que nos mandan informacion, esto nos quita tiempo y si pudieramos automatizar el envio de estas solicitudes de informacion seria algo muy beneficioso

### Solucion
Para solucionar estos dos problemas se crearon dos herramientas que tienen como objetivo eliminar por completo y acortar el tiempo que nos toma hacer lo que se menciono anteriormente, se creara una macro que junta archivos, y una herramienta en google script que mande correos automaticos, de tal forma no tendremos que juntar los archivos de uno por uno sino que se juntaran todos a la vez reduciendo asi el tiempo que nos toma limpiar la informacion y la herramienta de correos automaticos basicamente hace que no tengamos que volver a mandar un correo de forma manual, sino que los manda en automatico desde la cuenta que crea el script ahorrandonos esa actividad por completo

### Arquitectura
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/591cbef1-761a-4811-831b-32de2e1fd52b)
La arquitectura de la solucion va a ser como se muestra en la imagen, en resumidas cuentas se tienen que seguir los pasos que se mencionan en las guias de uso para su configuracion y entonces se podra trabajar, en la macro se seleccionan los archivos y se juntan en uno solo, lo que realmente esta pasando es que la macro une todos los archivos en hojas diferentes y luego se acciona un sub diferente que junta todos esas hojas en una sola y de esa forma se juntan todos los archivos, en el caso de los correos se crea un codigo que al introducir los datos de asunto, destinatario y copias se va a crear un correo electronico el cual llama una funcion para su envio, la forma en la que esto se automatiza es condfigurando activadores que se activen segun la temporalidad establecida 

# Tabla de contenidos
Esta se encuentra dandole clic al boton que se encuentra en la imagen, configure la pagina de github para que se puedan visualizar de la siguiente forma: 
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/0d68fef0-4e4d-4551-87a5-76fc1aab2d41)
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/8457ceb2-823c-4397-b2f7-bf491037d300)

# Requerimientos
los requerimientos que requieren las dos herramientas son los siguientes:
### Macro
Microsoft excel 2007 en adelante
### Correos automaticos
Cuenta de google

Debido a la forma en la que se trabajo y en las necesidades del area en donde laboro no habia necesidad de una aplicacion compleja que necesitara de servidores o entornos de desarrollo, lo que se busca son herramientas de facil acceso y configuracion y eso fue lo que se creo

# Instalacion
Una alternativa a una configuracion manual como se va a demostrar adelante es descargar las macros desde el enlace mostrado y solo seguir el uso descrito mas adelante

### Correos automaticos
Siguiendo los pasos descritos en la guia de uso, lamentablemente al ser un proyecto en la nube es imposible compartir un ejecutable o archivo que facilite su despliegue
### Macro que une archivos
Acceder al siguiente enlace: https://drive.google.com/drive/folders/1ojDgEdt8PgaAPGmLJc_TFqUC074cvF86 
Descargar los tres archivos los cuales son las macros necesarias para unir cada tipo de archivo segun su nombre, ejemplo, el archivo csv junta archivos de tipo csv, etc.
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/f5787968-a067-4f6e-82bc-3875f6fcf6b0)

Si aparece este mensaje al abrir el archivo dar clic en habilitar para permitir su uso, despues seguir instrucciones descritas adelante
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/c31c40df-976d-44a7-942b-b627c328838b)


# Guias de uso / Configuracion
### Correos Automaticos
A continuacion se mostrara los pasos para instalar e implementar la aplicacion de google script para automatizar el envio de correos

Requerimientos:

Cuenta de google

Pasos:

1. Iniciar sesion de Google Script
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/36f81951-79b9-4ecb-8a76-ab0321931973)

2. Abrir un nuevo proyecto
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/60ef3799-554d-4caf-8ec4-8064902daad7)

3. Copiar y pegar el codigo de Correos_Automaticos en el repositorio de GitHub y hacer los cambios necesarios
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/bc13b72f-a135-4b92-94a9-6d7114e41baf)
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/d3d1fe93-aad9-479c-bb5e-81dc031f7ece)

4. Agregar y ajustar el activador segun el tiempo necesario

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/2574d988-9a58-4efc-9c5f-d3c14f82a26f)
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/62359206-0a67-4f00-9822-74d26a737071)
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/8ddee0fe-91fd-4d37-8e6c-37f8f711f6e8)
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/ccd9c97c-e145-4352-a09b-1e10045203c8)

Despues de este paso solo quedaria esperar a que se manden los correos de forma automatica

### Macro que une archivos

A continuacion demostrare todos los pasos para instalar y ejecutar el codigo de visualbasic para llevar a cabo el uso de la macro que junta archivos

Requerimientos: 

Microsoft Excel 2007 en adelante

Pasos:

1. Abrir excel y verificar que se tenga activa la ventana developer (pasos para activarlo en caso de que no)
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/e66ac3be-e5a3-4d45-9f43-a29748ee18fe)
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/87215e6c-b601-46d3-aa1f-ff24272ac5ae)
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/611cc5dd-1dfd-47b7-a3fd-a1269a38c56c)

2. Seleccionar en la pestaña programador (developer) la opcion visual basic
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/ea207347-480c-46bc-aac8-b3025b1e79f4)
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/fadb0796-a24a-4bd9-8716-e50c74a7ff48)

3. Insertar un modulo dandole clic derecho a la carpeta de la libreta
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/d74bb00d-2221-499a-b957-f102aafe8229)

4. Copiar y pegar codigo de la pagina de GitHub

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/59397658-d85f-4da1-a26f-8f907086972d)
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/11b7ec31-29d6-4c23-9c45-1763116ae991)
5. Guardar proyecto y cerrar pestaña
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/9fec784e-3340-4563-b765-8812a51b4b88)
6. Darle clic a alt+F8 para accionar macro y luego darle enter
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/5a94896c-7153-4283-8339-c0067bc63fd3)
7. Seleccionar archivos a juntar
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/c155d6cf-f696-462b-ac65-90ec23ac14bb)
8. Importante guardar el proyecto como macro enabled workbook para uso futuro
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/8780b855-29bc-4e47-85ce-daf923bc86c8)

Despues de ese paso deberia poder juntar los archivos seleccionados en una sola hoja

# Guia de contribucion
En caso de querer contribuir al proyecto favor de seguir los siguientes pasos:

1. Crear un nuevo branch, de esa forma se clona el repositorio
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/54ea1a96-4e93-42e3-acf7-43dfc135eb5d)
![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/8df76196-7440-409a-87c7-2be633f1cbbf)

Asegurarse que el source sea main (master en este repositorio)

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/b44424c8-5cb1-42d0-b94e-80d218b0e624)

2. Realizar cambios en la branch creada

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/79e5b09d-c2b2-43a3-9623-995855391bea)

En este ejemplo se va a editar el README del repositorio

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/8dd748a2-4895-4008-bdde-89e37a2fa867)

Asegurarse que el branch sea el creado y hacer el commit directamente en esa branch

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/2916f833-b4c6-4ea2-9a93-eae8080d472f)

3. Realizar el pull request

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/85d0b0b5-9f0c-44df-becd-e2de38cd74f1)

Al darle clic a nuevo pull request selecciona como base la branch main y en compare la creada

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/b7e5db47-675a-4dde-bb35-facf6af779eb)

Observar los cambios que suceden y darle clic a create pull request

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/eae280ad-3440-43a3-b78a-d1828570ee8a)

Puedes agregar una descripcion de los cambios que estas realizando

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/e4866f53-0290-4523-8ffc-9a64da4a3256)

4. Esperar a que se acepte el merge

Desde este paso es tarea de los administradores del repositorio el aceptar el pull request creado para combinarlo con el main

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/a2f14d6c-eaf6-4850-a232-0ad1808add9d)

Una vez se acepte los cambios se reflejaran al main (master)

![image](https://github.com/MonkeTimes/Herrmientas-BI/assets/144874541/c369069a-85ce-4c1f-8123-c3365e1683e0)

#  Roadmap
En un futuro se espera implementar lo siguiente en los codigos:

### Macro

1. Funcion que permita identificar el tipo de archivo a juntar quitando la necesidad de usar diferentes archivos o editar el codigo de la macro

### Correos automaticos

1. Forma de dar formato a los mensajes (implementando HTML)

2. Forma de juntar todos los envios de informacion en la misma funcion, de tal forma deja de ser necesario tener un activador para cada KAM
