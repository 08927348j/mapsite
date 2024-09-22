<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Buscar Coordenadas</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            text-align: center;
            margin-top: 50px;
            background-color: #f44336; /* Rojo */
        }

        h1 {
            color: white;
            font-family: 'Comic Sans MS', cursive, sans-serif; /* Fuente Comic */
        }

        button {
            padding: 10px 20px;
            font-size: 16px;
            margin: 10px;
            cursor: pointer;
            border: none;
            border-radius: 5px;
        }

        .buscar {
            background-color: #FF9800; /* Naranja */
            color: white;
        }

        .ingresar {
            background-color: #2196F3; /* Azul */
            color: white;
        }

        #mapButton {
            background-color: #4CAF50; /* Verde */
            color: white;
            display: none; /* Oculto inicialmente */
        }

        #searchContainer, #nuevoSiteContainer, #eliminarContainer, #excelContainer {
            display: none; /* Oculto inicialmente */
            margin-top: 20px;
        }

        #searchInput, #nuevaCadena, #siteAEliminar {
            padding: 10px;
            font-size: 16px;
            width: 200px;
            border: 2px solid #ccc;
            border-radius: 5px;
        }

        #searchButton, #eliminarButton, #refreshButton {
            padding: 10px 20px;
            font-size: 16px;
            margin-left: 10px;
            cursor: pointer;
            border: none;
            border-radius: 5px;
        }

        #searchButton {
            background-color: #4CAF50; /* Verde para buscar */
            color: white;
        }

        #result {
            margin-top: 20px;
            font-size: 18px;
            font-weight: bold;
        }

        #loadButton {
            background-color: #2196F3; /* Azul */
            color: white;
        }

        #refreshButton {
            background-color: #FFEB3B; /* Amarillo */
            color: #2196F3; /* Azul */
            display: none; /* Oculto inicialmente */
        }

        #successMessage {
            margin-top: 20px;
            color: white;
            font-size: 18px;
            display: none; /* Oculto inicialmente */
        }

        #loadingMessage {
            color: white;
            display: none; /* Oculto inicialmente */
        }
    </style>
</head>
<body>

    <h1>Buscar Coordenadas</h1>

    <button class="buscar" onclick="mostrarBarraBusqueda()">Ingresar búsqueda</button>
    <button class="ingresar" onclick="mostrarFormularioIngresar()">Ingresar nuevo Site</button>

    <div id="searchContainer">
        <input type="text" id="searchInput" maxlength="6" placeholder="Ingrese hasta 6 caracteres alfanuméricos">
        <button id="searchButton" onclick="buscarCoordenadas()">Buscar</button>
    </div>

    <div id="result"></div>

    <button id="mapButton" style="display: none;" onclick="abrirMapa()">Mapa</button>

    <div id="eliminarContainer" style="display:none;">
        <input type="text" id="siteAEliminar" placeholder="Site a eliminar">
        <button onclick="eliminarSite()">Eliminar</button>
    </div>

    <div id="nuevoSiteContainer" style="display:none;">
        <h2>Ingresar nuevo Site</h2>
        <input type="text" id="nuevaCadena" maxlength="6" placeholder="Nueva cadena alfanumérica">
        <input type="text" id="nuevasCoordenadas" placeholder="Coordenadas">
        <button onclick="agregarNuevoSite()">Guardar</button>
    </div>

    <div id="excelContainer" style="display:none;">
        <h2>Cargar archivo WRITE</h2>
        <input type="file" id="inputExcel" accept=".xlsx, .xls">
        <button id="loadButton" onclick="procesarExcel()">Cargar</button>
    </div>

    <div id="loadingMessage">Cargando...</div>
    <div id="successMessage">Carga exitosa de datos desde el archivo WRITE.</div>
    <button id="refreshButton" onclick="refreshPage()">Refrescar</button>

    <script>
        let database = {
            'ABC123': '40.7128° N, 74.0060° W', // Coordenadas de Nueva York
            'XYZ456': '34.0522° N, 118.2437° W', // Coordenadas de Los Ángeles
            'DELETE': '51.5074° N, 0.1278° W' // Coordenadas de Londres (ejemplo)
        };

        function mostrarBarraBusqueda() {
            document.getElementById("searchContainer").style.display = "block";
        }

        function mostrarFormularioIngresar() {
            document.getElementById("nuevoSiteContainer").style.display = "block";
        }

        function buscarCoordenadas() {
            const cadenaAlfanumerica = document.getElementById("searchInput").value.toUpperCase(); // Convertir a mayúsculas

            if (cadenaAlfanumerica.length === 0) {
                alert("Por favor, ingrese una cadena alfanumérica.");
                return;
            }

            if (cadenaAlfanumerica === "WRITE") {
                generarExcel();
                return;
            }

            if (cadenaAlfanumerica === "READ") {
                document.getElementById("excelContainer").style.display = "block";
                return;
            }

            if (cadenaAlfanumerica === "BORRAR") {
                document.getElementById("eliminarContainer").style.display = "block"; // Mostrar contenedor de eliminar
                document.getElementById("result").textContent = ''; // Limpiar resultado
                document.getElementById("mapButton").style.display = "none"; // Ocultar botón de mapa
                return;
            }

            if (database[cadenaAlfanumerica]) {
                const coordenadas = database[cadenaAlfanumerica];
                document.getElementById("result").textContent = `Coordenadas encontradas: ${coordenadas}`;
                document.getElementById("mapButton").style.display = "inline-block";
                document.getElementById("eliminarContainer").style.display = "none"; // Asegurar que no se muestre contenedor de eliminar
            } else {
                document.getElementById("result").textContent = 'Coordenadas no encontradas';
                document.getElementById("mapButton").style.display = "none"; // Ocultar botón de mapa si no se encuentra
                document.getElementById("eliminarContainer").style.display = "none"; // Ocultar contenedor de eliminar
            }
        }

        function agregarNuevoSite() {
            const nuevaCadena = document.getElementById("nuevaCadena").value.toUpperCase(); // Convertir a mayúsculas
            const nuevasCoordenadas = document.getElementById("nuevasCoordenadas").value;

            if (nuevaCadena.length === 0 || nuevasCoordenadas.length === 0) {
                alert("Por favor, complete ambos campos.");
                return;
            }

            if (database[nuevaCadena]) {
                alert('La cadena ya existe en la base de datos');
                return;
            }

            database[nuevaCadena] = nuevasCoordenadas;
            alert("Nuevo site agregado con éxito.");

            document.getElementById("nuevaCadena").value = "";
            document.getElementById("nuevasCoordenadas").value = "";
            document.getElementById("nuevoSiteContainer").style.display = "none"; // Ocultar formulario
        }

        function eliminarSite() {
            const siteAEliminar = document.getElementById("siteAEliminar").value.toUpperCase();

            if (siteAEliminar.length === 0) {
                alert("Por favor, ingrese el nombre del site a eliminar.");
                return;
            }

            if (database[siteAEliminar]) {
                delete database[siteAEliminar];
                alert(`Site ${siteAEliminar} eliminado con éxito.`);
                document.getElementById("siteAEliminar").value = ""; // Limpiar el campo
            } else {
                alert("No se encontró el site para eliminar.");
            }
        }

        function abrirMapa() {
            const coordenadas = document.getElementById("result").textContent.match(/[-\d.]+/g);
            if (coordenadas) {
                const lat = coordenadas[0];
                const lng = coordenadas[1];
                const url = `https://www.google.com/maps?q=${lat},${lng}`;
                window.open(url, '_blank');
            }
        }

        function generarExcel() {
            const wb = XLSX.utils.book_new();
            const wsData = [["Site", "Coordenadas"]];

            for (const [key, value] of Object.entries(database)) {
                wsData.push([key, value]);
            }

            const ws = XLSX.utils.aoa_to_sheet(wsData);
            XLSX.utils.book_append_sheet(wb, ws, "Sites");

            XLSX.writeFile(wb, "coordenadas_sites.xlsx");
        }

        function procesarExcel() {
            const fileInput = document.getElementById('inputExcel');
            const file = fileInput.files[0];

            if (!file) {
                alert('Por favor, seleccione un archivo WRITE.');
                return;
            }

            const reader = new FileReader();

            // Mostrar mensaje de carga
            document.getElementById('loadingMessage').style.display = 'block';

            reader.onload = function (e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

                let validData = true;
                jsonData.slice(1).forEach(row => {
                    const site = row[0];
                    const coordenadas = row[1];

                    if (!site || !coordenadas) {
                        validData = false; // Marca como inválido si hay datos faltantes
                    } else {
                        database[site.toUpperCase()] = coordenadas;
                    }
                });

                // Ocultar mensaje de carga
                document.getElementById('loadingMessage').style.display = 'none';

                if (!validData) {
                    alert('Algunos datos en el archivo son inválidos. Asegúrate de que todas las filas tengan un site y coordenadas.');
                    return;
                }

                document.getElementById('successMessage').style.display = 'block';
                document.getElementById('refreshButton').style.display = 'inline-block';
                document.getElementById('excelContainer').style.display = 'none';
            };

            reader.readAsArrayBuffer(file);
        }

        function refreshPage() {
            location.reload();
        }
    </script>

</body>
</html>
