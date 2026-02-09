<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Informe FeNO - Formulario Manual</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            padding: 20px;
            border: 1px solid #ccc;
            max-width: 1000px;
        }
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 30px;
        }
        .header img {
            height: 60px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }
        table, th, td {
            border: 1px solid black;
        }
        th, td {
            padding: 8px;
            text-align: left;
        }
        .section-title {
            background-color: #f2f2f2;
            font-weight: bold;
            padding: 10px;
            margin-top: 20px;
        }
        .curva-section {
            display: flex;
            justify-content: space-between;
            margin: 20px 0;
        }
        .curva-box {
            border: 1px solid #000;
            padding: 10px;
            width: 48%;
            height: 200px;
            display: flex;
            align-items: center;
            justify-content: center;
            background-color: #f9f9f9;
        }
        .result-box {
            border: 1px solid #000;
            padding: 20px;
            text-align: center;
            margin: 20px 0;
            background-color: #f0f8ff;
        }
        .references {
            font-size: 12px;
            margin-top: 30px;
            padding-top: 10px;
            border-top: 1px solid #ccc;
        }
        input[type="text"] {
            border: none;
            border-bottom: 1px solid #000;
            width: 100%;
            padding: 5px;
            box-sizing: border-box;
        }
    </style>
</head>
<body>

    <div class="header">
        <img src="placeholder_logo1.png" alt="Logo INT">
        <img src="placeholder_logo2.png" alt="Logo Otro">
    </div>

    <table>
        <tr>
            <td>Nombre:</td>
            <td><input type="text" id="nombre" placeholder="Ingrese nombre"></td>
            <td>Apellidos:</td>
            <td><input type="text" id="apellidos" placeholder="Ingrese apellidos"></td>
        </tr>
        <tr>
            <td>RUT:</td>
            <td><input type="text" id="rut" placeholder="Ej: 12345678-9"></td>
            <td>Género:</td>
            <td><input type="text" id="genero" placeholder="Ej: Femenino"></td>
        </tr>
        <tr>
            <td>Operador:</td>
            <td><input type="text" id="operador" placeholder="Ej: Klgo. Christian Sáez"></td>
            <td>Médico:</td>
            <td><input type="text" id="medico" placeholder="Ej: Dra Patricia Schonffeldt"></td>
        </tr>
        <tr>
            <td>F. nacimiento:</td>
            <td><input type="text" id="fnac" placeholder="DD/MM/AAAA"></td>
            <td>Edad:</td>
            <td><input type="text" id="edad" placeholder="Ej: 59"></td>
        </tr>
        <tr>
            <td>Altura:</td>
            <td><input type="text" id="altura" placeholder="Ej: 166"></td>
            <td>Peso:</td>
            <td><input type="text" id="peso" placeholder="Ej: 90"></td>
        </tr>
        <tr>
            <td>Raza:</td>
            <td><input type="text" id="raza" placeholder="Ej: Caucásica"></td>
            <td>Procedencia:</td>
            <td><input type="text" id="procedencia" placeholder="Ej: Poli"></td>
        </tr>
        <tr>
            <td>Fecha de Examen:</td>
            <td><input type="text" id="fecha_examen" placeholder="DD/MM/AAAA"></td>
            <td></td>
            <td></td>
        </tr>
    </table>

    <div class="section-title">Prueba de Óxido Nítrico Exhalado</div>
    
    <p><strong>Predictivos:</strong> ATS/ERS <strong>Equipo:</strong> CA2122 FeNO (Sunvou)</p>

    <table>
        <tr>
            <td>Temperatura:</td>
            <td><input type="text" id="temp" placeholder="Ej: 22.4 °C"></td>
        </tr>
        <tr>
            <td>Presión:</td>
            <td><input type="text" id="presion" placeholder="Ej: 13.3 cmH2O"></td>
        </tr>
        <tr>
            <td>Tasa de Flujo:</td>
            <td><input type="text" id="flujo" placeholder="Ej: 52 ml/s"></td>
        </tr>
    </table>

    <div class="section-title">Curva de Exhalación y Análisis</div>
    
    <div class="curva-section">
        <div class="curva-box">
            <p><strong>Espacio para Curva de Exhalación</strong><br>
            (Pegar imagen aquí)</p>
        </div>
        <div class="curva-box">
            <p><strong>Espacio para Análisis de Curva</strong><br>
            (Pegar imagen aquí)</p>
        </div>
    </div>

    <div class="result-box">
        <h2>FeNO<sub>50</sub>: <input type="text" id="feno" placeholder="Ej: 38 ppb" style="width: 100px; text-align: center;"></h2>
    </div>

    <div class="references">
        <p><strong>Referencias:</strong><br>
        Dweik RA, Boggs PB, Erzurum SC, et al. An official ATS clinical practice guideline: interpretation of exhaled nitric oxide levels (FENO) for clinical applications. <em>Am J Respir Crit Care Med</em>. 2011;184(5):602-615. doi:10.1164/rccm.9120-11ST</p>
    </div>

</body>
</html>
