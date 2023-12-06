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


# Guias de uso / Instalacion
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

Despues de ese paso deberia poder juntar los archivos seleccionados en una sola hoja






