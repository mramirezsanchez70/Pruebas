/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { DatosXML } from "./ExtraerXML";
import { validarArchivos } from "./ValidadorArchivos";
import { mostrarModal } from "./UI";

export function CargaXML_PDF() {

  const inputFile = crearInputArchivo();
  inputFile.value = "";
  inputFile.click();

  inputFile.onchange = async () => {

    const files = inputFile.files;

    if (!files || files.length === 0) {
      mostrarModal("No seleccionaste archivos");
      return;
    }

    const resultado = validarArchivos(files);

    if (resultado.error) {
      mostrarModal(resultado.error);
      inputFile.value = "";
      return;
    }

    if (resultado.xmlFile) {
      await DatosXML(resultado.xmlFile);
    }

    inputFile.value = "";
  };
}