/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// 🔹 Se ejecuta cuando Office está listo
Office.onReady((info) => {

  // ✅ 1.- Validamos que estamos en Excel
  if (info.host === Office.HostType.Excel) {

    ocultarSideload(); // 🔹 Oculta pantalla inicial

    const inputFile = crearInputArchivo(); // 🔹 Input oculto

    const button = document.getElementById("btnCarga");

    // ✅ 1.- CLICK EN BOTÓN → ABRE SELECTOR DE ARCHIVOS
    if (button) {
      button.addEventListener("click", () => {
        inputFile.value = ""; // 🔥 Limpia selección anterior
        inputFile.click();    // 🔹 Abre ventana para elegir archivos
      });
    }

    // ✅ 2.- CUANDO EL USUARIO DA CLICK EN "ABRIR"
    inputFile.addEventListener("change", async () => {

      const files = inputFile.files;

      // 🔹 Validación básica
      if (!files || files.length === 0) {
        mostrarModal("No seleccionaste archivos, debes seleccionar un xml con su pdf para su comparativa");
        return;
      }

      const archivos = Array.from(files);

      // 🔹 Contadores
      let contadorXML = 0;
      let contadorPDF = 0;

      let xmlFile = null;
      let pdfFile = null;

      // 🔹 Detecta tipo de archivos
      for (const file of archivos) {

        const ext = file.name.split(".").pop()?.toLowerCase();

        if (ext === "xml") {
          contadorXML++;
          xmlFile = file;
        } else if (ext === "pdf") {
          contadorPDF++;
          pdfFile = file;
        }
      }

      // ✅ 2b.- MÁS DE 2 ARCHIVOS
      if (archivos.length > 2) {
        mostrarModal("Selección incorrecta: para su comparativa, debe seleccionar solo 2 archivo, un xml y su pdf");
        inputFile.value = "";
        return;
      }

      // ✅ 2c.- SOLO 1 ARCHIVO
      if (archivos.length === 1) {
        mostrarModal("Debe seleccionar un xml y su pdf para su comparativa");
        inputFile.value = "";
        return;
      }

      // ✅ 2a.- EXACTAMENTE 2 ARCHIVOS
      if (archivos.length === 2) {

        // ✅ Validar si son del mismo tipo
        if (contadorXML === 2 || contadorPDF === 2) {
          mostrarModal("Selección incorrecta: para su comparativa, debe seleccionar solo el xml con su pdf");
          inputFile.value = "";
          return;
        }

        // ✅ Validar combinación correcta (1 XML + 1 PDF)
        if (contadorXML === 1 && contadorPDF === 1) {

          // 🔥 ESCRIBE EN EXCEL (CASO CORRECTO)
          await escribirEnExcel(xmlFile.name, pdfFile.name);

          inputFile.value = "";
          return;
        }
      }
    });
  }
});


/**
 * 🔹 Muestra modal con mensaje
 */
function mostrarModal(mensaje) {

  const modal = document.getElementById("modalError");
  const texto = document.getElementById("modalMensaje");
  const btnCerrar = document.getElementById("btnCerrarModal");

  texto.textContent = mensaje;
  modal.style.display = "block";

  // 🔹 Cerrar modal
  btnCerrar.onclick = () => {
    modal.style.display = "none";
  };
}


/**
 * 🔹 Crea input tipo file oculto
 */
function crearInputArchivo() {

  let input = document.getElementById("fileInput");

  if (!input) {
    input = document.createElement("input");
    input.type = "file";
    input.accept = ".xml,.pdf"; // ✅ SOLO XML Y PDF
    input.multiple = true;      // ✅ PERMITE SELECCIONAR VARIOS
    input.style.display = "none";
    input.id = "fileInput";

    document.body.appendChild(input);
  }

  return input;
}


/**
 * 🔹 Oculta pantalla inicial de Office
 */
function ocultarSideload() {

  const sideload = document.getElementById("sideload-msg");
  if (sideload) sideload.style.display = "none";

  const appBody = document.getElementById("app-body");
  if (appBody) appBody.style.display = "flex";
}


/**
 * 🔥 Escribe en Excel
 */
async function escribirEnExcel(xmlNombre, pdfNombre) {

  await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const texto = `${xmlNombre} | ${pdfNombre}`;

    const range = sheet.getRange("A1");
    range.values = [[texto]];

    sheet.activate();

    await context.sync();
  });
}