import { Injectable } from '@angular/core';
import * as JSZip from 'jszip';

export type IntervalosResultado = {
  [coluna: string]: {
    [linha: number]: string | number | null;
  };
};

@Injectable({
  providedIn: 'root'
})
export class ExcelReaderService {

  // ======================================================
  // API PÚBLICA
  // ======================================================
  async lerIntervalos(
    file: ArrayBuffer,
    aba: number,              // 1 = primeira aba, 2 = segunda...
    intervalos: string[]      // ex: ["H12-H300", "I10-I20", "H2"]
  ): Promise<IntervalosResultado> {

    const zip = await JSZip.loadAsync(file);

    const sheetPath = `xl/worksheets/sheet${aba}.xml`;
    const sheetXml = await zip.file(sheetPath)?.async('string');

    if (!sheetXml) {
      throw new Error(`Aba ${aba} não encontrada (${sheetPath})`);
    }

    const sharedStringsXml = await zip
      .file('xl/sharedStrings.xml')!
      .async('string');

    const sharedStrings = this.parseSharedStrings(sharedStringsXml);

    const resultado: IntervalosResultado = {};

    for (const intervalo of intervalos) {
      const { col, start, end } = this.parseInterval(intervalo);

      // 🔧 NÃO sobrescreve mais colunas existentes
      if (!resultado[col]) {
        resultado[col] = {};
      }

      for (let row = start; row <= end; row++) {
        const ref = `${col}${row}`;
        resultado[col][row] = this.getCellValue(
          sheetXml,
          ref,
          sharedStrings
        );
      }
    }

    return resultado;
  }

  // ======================================================
  // LÊ UMA CÉLULA
  // ======================================================
  private getCellValue(
    sheetXml: string,
    cellRef: string,
    sharedStrings: string[],
    styles: number[]
  ): string | number | null {

    const cellRegex = new RegExp(
      `<c[^>]*r="${cellRef}"([^>]*)>([\\s\\S]*?)<\\/c>`
    );

    const match = sheetXml.match(cellRegex);
    if (!match) return null;

    const attrs = match[1];
    const innerXml = match[2];

    // Shared string
    if (/t="s"/.test(attrs)) {
      const v = innerXml.match(/<v>(\d+)<\/v>/);
      return v ? sharedStrings[Number(v[1])] ?? null : null;
    }

    const v = innerXml.match(/<v>([\s\S]*?)<\/v>/);
    if (!v) return null;

    const raw = Number(v[1]);
    if (isNaN(raw)) return v[1];

    // 🎯 Estilo da célula
    const styleMatch = attrs.match(/s="(\d+)"/);
    if (!styleMatch) return raw;

    const styleIndex = Number(styleMatch[1]);
    const numFmtId = styles[styleIndex];

    // 📅 DATA
    if ([14, 15, 16, 17, 22].includes(numFmtId)) {
      return this.excelDateToString(raw);
    }

    // 📊 PERCENTUAL
    if ([9, 10].includes(numFmtId)) {
      return raw;
    }

    // 💰 MOEDA / NÚMERO
    return raw;
  }


  // ======================================================
  // PARSE sharedStrings.xml
  // ======================================================
  private parseSharedStrings(xml: string): string[] {
    const result: string[] = [];
    const regex = /<si>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/si>/g;

    let match: RegExpExecArray | null;
    while ((match = regex.exec(xml))) {
      result.push(match[1]);
    }

    return result;
  }

  // ======================================================
  // PARSE "H12-H300" ou "H2"
  // ======================================================
  private parseInterval(interval: string): {
    col: string;
    start: number;
    end: number;
  } {

    // Caso 1: intervalo (ex: H12-H300)
    const rangeMatch = interval.match(/^([A-Z]+)(\d+)-\1(\d+)$/i);
    if (rangeMatch) {
      return {
        col: rangeMatch[1].toUpperCase(),
        start: Number(rangeMatch[2]),
        end: Number(rangeMatch[3])
      };
    }

    // Caso 2: célula única (ex: H2)
    const singleMatch = interval.match(/^([A-Z]+)(\d+)$/i);
    if (singleMatch) {
      const row = Number(singleMatch[2]);
      return {
        col: singleMatch[1].toUpperCase(),
        start: row,
        end: row
      };
    }

    throw new Error(`Intervalo inválido: ${interval}`);
  }

  // ======================================================
  // CONVERTE DATA DO EXCEL (1900-based) → dd/MM/yyyy
  // ======================================================
  private excelDateToString(serial: number): string {
    const excelEpoch = new Date(1899, 11, 30); // Excel base
    const date = new Date(excelEpoch.getTime() + serial * 86400000);

    const dd = String(date.getDate()).padStart(2, '0');
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const yyyy = date.getFullYear();

    return `${dd}/${mm}/${yyyy}`;
  }
}
