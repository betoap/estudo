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
export class ExcelRangeReaderService {

  // ======================================================
  // API PÚBLICA
  // ======================================================
  async lerIntervalos(
    file: ArrayBuffer,
    aba: number,              // 1 = primeira aba, 2 = segunda...
    intervalos: string[]      // ex: ["H12-H300", "I10-I20"]
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

      resultado[col] = {};

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
    sharedStrings: string[]
  ): string | number | null {

    const cellRegex = new RegExp(
      `<c[^>]*r="${cellRef}"[^>]*>([\\s\\S]*?)<\\/c>`
    );

    const match = sheetXml.match(cellRegex);
    if (!match) return null;

    const fullCell = match[0];
    const innerXml = match[1];

    // Texto compartilhado
    if (/t="s"/.test(fullCell)) {
      const v = innerXml.match(/<v>(\d+)<\/v>/);
      return v ? sharedStrings[Number(v[1])] ?? null : null;
    }

    // Número / fórmula
    const v = innerXml.match(/<v>([\s\S]*?)<\/v>/);
    if (!v) return null;

    const raw = v[1];
    const num = Number(raw);

    return isNaN(num) ? raw : num;
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
  // PARSE "H12-H300"
  // ======================================================
  private parseInterval(interval: string): {
    col: string;
    start: number;
    end: number;
  } {

    const match = interval.match(/^([A-Z]+)(\d+)-\1(\d+)$/i);

    if (!match) {
      throw new Error(`Intervalo inválido: ${interval}`);
    }

    return {
      col: match[1].toUpperCase(),
      start: Number(match[2]),
      end: Number(match[3])
    };
  }
}
// USO
constructor(private excel: ExcelRangeReaderService) {}

async carregar(file: File) {
  const buffer = await file.arrayBuffer();

  const dados = await this.excel.lerIntervalos(
    buffer,
    2,
    ['H12-H300', 'I10-I20', 'K1-K50']
  );

  console.log(dados);
}

/*

Exemplo de retorno
{
  H: {
    12: 1300,
    13: 1400,
    14: null
  },
  I: {
    10: "07/07/2026",
    11: "07/08/2026"
  },
  K: {
    1: 999,
    2: 888
  }
}

*/

// ======================================================
// PARSE "H12-H300" OU "H12"
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




for (const intervalo of intervalos) {
  const { col, start, end } = this.parseInterval(intervalo);

  // 🔧 NÃO sobrescreve mais
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
