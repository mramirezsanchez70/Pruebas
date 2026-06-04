/*
 * Copyright (c) Microsoft Corporation. Licensed under MIT
 */

/* global console, document, Excel, Office */

///////////////////////////////////////////////////////////////////////////
// 🔥 FUNCIÓN PRINCIPAL
// 👉 Lee XML, valida UUID, gestiona duplicados y escribe en Excel
///////////////////////////////////////////////////////////////////////////
export async function DatosXML(file: File) {

  try {

    // ✅ 1.- Leer archivo XML como texto
    // 👉 Convierte el archivo recibido en un string para poder procesarlo
    const textoXML = await leerArchivo(file);


    // ✅ 2.- Parsear XML a objeto DOM
    // 👉 Convierte el texto en una estructura manipulable tipo HTML/XML
    const parser = new DOMParser();
    const xml = parser.parseFromString(textoXML, "text/xml");

    // ✅ 3.- Validar si el XML tiene errores de estructura
    // 👉 Si el parser detecta errores, detiene el proceso
    if (xml.getElementsByTagName("parsererror").length > 0) {
      throw new Error("El archivo XML no es válido");
    }

    // ✅ 4.- Obtener nodos principales del CFDI
    // 👉 Se extraen las secciones clave del XML fiscal
    const comprobante = xml.getElementsByTagName("cfdi:Comprobante")[0];
    const emisor = xml.getElementsByTagName("cfdi:Emisor")[0];
    const receptor = xml.getElementsByTagName("cfdi:Receptor")[0];
    const timbres = xml.getElementsByTagName("tfd:TimbreFiscalDigital");

    // ✅ 5.- Extraer datos fiscales principales
    // 👉 Se guardan valores importantes en un objeto
    const datos = {
      rfcEmisor: emisor?.getAttribute("Rfc") || "",
      nombreEmisor: emisor?.getAttribute("Nombre") || "",
      rfcReceptor: receptor?.getAttribute("Rfc") || "",
      nombreReceptor: receptor?.getAttribute("Nombre") || "",
      total: comprobante?.getAttribute("Total") || "",
      fecha: comprobante?.getAttribute("Fecha") || "",
      tipo: comprobante?.getAttribute("TipoDeComprobante") || "",
      uuids: [] as string[]
    };

    // ✅ 6.- Extraer TODOS los UUID del XML
    // 👉 Un XML puede tener más de un UUID
    for (let i = 0; i < timbres.length; i++) {
      const uuid = timbres[i].getAttribute("UUID");
      if (uuid) {
        datos.uuids.push(uuid);
      }
    }

    // ✅ 7.- Validar que exista al menos un UUID
    // 👉 Si no hay UUID, el XML no es usable
    if (datos.uuids.length === 0) {
      throw new Error("El XML no contiene UUID");
    }

    // ✅ 8.- Validar UUIDs contra Hoja2
    // 👉 Verifica si ya existen en la base de Excel
    const existeUUID = await validarUUIDs(datos.uuids);

    // ✅ 9.- Manejar duplicados
    // 👉 Si existen UUID repetidos, mostrar modal
    if (existeUUID.length > 0) {

      const continuar = await mostrarModalUUID(existeUUID);

      // ✅ 10.- Si usuario cancela, detener ejecución
      if (!continuar) {
        return;
      }

      // ✅ 11.- Si usuario acepta, crear devolución
      await crearDevolucion(datos);
    }

    // ✅ 12.- Escribir datos en Excel
    // 👉 Inserta información en la hoja activa
    await escribirDatosExcel(datos);

    return datos;

  } catch (error) {

    // ✅ 13.- Manejo de errores
    // 👉 Captura y muestra cualquier fallo
    console.error("❌ Error en DatosXML:", error);

    const texto = (error as Error).message || "Error desconocido";
    alert(texto);
  }
}

///////////////////////////////////////////////////////////////////////////
// 🔹 FUNCIÓN: leerArchivo
///////////////////////////////////////////////////////////////////////////
function leerArchivo(file: File): Promise<string> {

  // ✅ 14.- Convertir archivo en texto
  // 👉 Usa FileReader para leer contenido del XML
  return new Promise((resolve, reject) => {

    const reader = new FileReader();

    reader.onload = () => resolve(reader.result as string);

    reader.onerror = () => reject("Error leyendo archivo");

    reader.readAsText(file);
  });
}

///////////////////////////////////////////////////////////////////////////
// 🔹 FUNCIÓN: validarUUIDs
///////////////////////////////////////////////////////////////////////////
async function validarUUIDs(uuids: string[]) {

  // ✅ 15.- Ejecutar en contexto de Excel
  // 👉 Permite interactuar con hojas de Excel
  return await Excel.run(async (context) => {

    const hoja = context.workbook.worksheets.getItem("Hoja2");
    const rango = hoja.getRange("B:B");

    rango.load("values");
    await context.sync();

    // ✅ 16.- Convertir columna en Set
    // 👉 Mejora rendimiento en búsqueda de UUID
    const valores = new Set(
      rango.values
        .map(f => f[0])
        .filter(v => v !== null && v !== "")
    );

    const existentes: string[] = [];

    // ✅ 17.- Comparar UUIDs del XML contra Excel
    for (const uuid of uuids) {
      if (valores.has(uuid)) {
        existentes.push(uuid);
      }
    }

    return existentes;
  });
}

///////////////////////////////////////////////////////////////////////////
// 🔹 FUNCIÓN: mostrarModalUUID
///////////////////////////////////////////////////////////////////////////
function mostrarModalUUID(uuids: string[]): Promise<boolean> {

  // ✅ 18.- Mostrar ventana modal al usuario
  // 👉 Permite decidir si continuar o no
  return new Promise((resolve) => {

    const modal = document.getElementById("modalError");
    const texto = document.getElementById("modalMensaje");

    const btnSi = document.getElementById("btnSi");
    const btnNo = document.getElementById("btnNo");

    texto.textContent = `UUID ya existe: ${uuids.join(", ")} ¿Deseas crear devolución?`;
    modal.style.display = "block";

    // ✅ 19.- Acción botón SI
    btnSi?.addEventListener("click", () => {
      modal.style.display = "none";
      resolve(true);
    }, { once: true });

    // ✅ 20.- Acción botón NO
    btnNo?.addEventListener("click", () => {
      modal.style.display = "none";
      resolve(false);
    }, { once: true });
  });
}

///////////////////////////////////////////////////////////////////////////
// 🔹 FUNCIÓN: crearDevolucion
///////////////////////////////////////////////////////////////////////////
async function crearDevolucion(datos: any) {

  // ✅ 21.- Lógica de devolución
  // 👉 Aquí puedes integrar procesos contables
  console.log("✅ Creando devolución...", datos);

  await Excel.run(async (context) => {

    const hoja = context.workbook.worksheets.getItem("Devoluciones");

    const fila = [
      datos.rfcEmisor,
      datos.total,
      datos.uuids.join(" | "),
      new Date().toLocaleString()
    ];

    // ✅ 22.- Insertar fila en hoja Devoluciones
    hoja.getRange("A1").getOffsetRange(1, 0).setValues([fila]);

    await context.sync();
  });
}

///////////////////////////////////////////////////////////////////////////
// 🔹 FUNCIÓN: escribirDatosExcel
///////////////////////////////////////////////////////////////////////////
async function escribirDatosExcel(datos: any) {

  // ✅ 23.- Insertar datos fiscales en Excel
  await Excel.run(async (context) => {

    const sheet = context.workbook.getActiveWorksheet();

    const fila = [
      datos.rfcEmisor,
      datos.nombreEmisor,
      datos.rfcReceptor,
      datos.nombreReceptor,
      datos.total,
      datos.fecha,
      datos.tipo,
      datos.uuids.join(" | ")
    ];

    const range = sheet.getRange("A3:H3");

    // ✅ 24.- Escribir valores
    range.setValues([fila]);

    // ✅ 25.- Crear encabezados
    const header = sheet.getRange("A2:H2");
    header.setValues([[
      "RFC Emisor", "Nombre Emisor",
      "RFC Receptor", "Nombre Receptor",
      "Total", "Fecha", "Tipo", "UUID"
    ]]);

    // ✅ 26.- Aplicar formato a encabezados
    header.getFormat().getFont().setBold(true);

    await context.sync();
  });
}