import { Injectable } from '@angular/core';
import * as JSZip from 'jszip';

export type IntervalosResultado = {
  [coluna: string]: {
    [linha: number]: string | number | null;
  };
};

type StylesInfo = {
  cellXfsNumFmtIds: number[];              // index = s="N" da célula
  numFmts: Record<number, string>;         // numFmtId -> formatCode (custom)
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

    // ✅ NOVO: estilos (para detectar data, percentual, etc.)
    const stylesXml = await zip.file('xl/styles.xml')?.async('string');
    const styles = stylesXml ? this.parseStyles(stylesXml) : {
      cellXfsNumFmtIds: [],
      numFmts: {}
    };

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
          sharedStrings,
          styles
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
    styles: StylesInfo
  ): string | number | null {

    const cellRegex = new RegExp(
      `<c[^>]*r="${cellRef}"([^>]*)>([\\s\\S]*?)<\\/c>`
    );

    const match = sheetXml.match(cellRegex);
    if (!match) return null;

    const attrs = match[1] ?? '';
    const innerXml = match[2] ?? '';

    // Shared string
    if (/t="s"/.test(attrs)) {
      const v = innerXml.match(/<v>(\d+)<\/v>/);
      return v ? (sharedStrings[Number(v[1])] ?? null) : null;
    }

    // Se não tem <v>, pode ser célula vazia ou fórmula sem cache
    const v = innerXml.match(/<v>([\s\S]*?)<\/v>/);
    if (!v) return null;

    const rawText = v[1];

    // Tenta número
    const rawNumber = Number(rawText);
    const isNumber = !Number.isNaN(rawNumber);

    // Se não é número, devolve texto
    if (!isNumber) return rawText;

    // 🎯 Estilo da célula (s="N")
    const styleMatch = attrs.match(/s="(\d+)"/);
    if (!styleMatch) return rawNumber;

    const styleIndex = Number(styleMatch[1]);
    const numFmtId = styles.cellXfsNumFmtIds[styleIndex];

    // Se não encontrou numFmtId (estilo fora do range), devolve o número
    if (numFmtId == null) return rawNumber;

    // ✅ Detecta formato customizado pelo formatCode
    const formatCode = styles.numFmts[numFmtId];

    // 📅 DATA (built-in comuns + custom com d/m/y)
    if (this.isDateFormat(numFmtId, formatCode)) {
      return this.excelDateToString(rawNumber);
    }

    // 📊 PERCENTUAL
    if (this.isPercentFormat(numFmtId, formatCode)) {
      // Excel guarda 1% como 0.01 (isso é correto).
      // Se você quiser retornar já "1" ao invés de 0.01, descomente a linha abaixo:
      // return rawNumber * 100;
      return rawNumber;
    }

    // 💰 MOEDA / NÚMERO
    return rawNumber;
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
  // PARSE styles.xml (cellXfs + numFmts)
  // ======================================================
  private parseStyles(xml: string): StylesInfo {
    // 1) numFmts (custom)
    const numFmts: Record<number, string> = {};

    const numFmtsBlock = xml.match(/<numFmts[^>]*>[\s\S]*?<\/numFmts>/);
    if (numFmtsBlock) {
      const fmtRegex = /<numFmt[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g;
      let m: RegExpExecArray | null;
      while ((m = fmtRegex.exec(numFmtsBlock[0]))) {
        const id = Number(m[1]);
        const code = m[2];
        numFmts[id] = code;
      }
    }

    // 2) cellXfs (isso aqui é o que o s="N" referencia)
    const cellXfsBlock = xml.match(/<cellXfs[^>]*>[\s\S]*?<\/cellXfs>/);
    if (!cellXfsBlock) {
      return { cellXfsNumFmtIds: [], numFmts };
    }

    const cellXfsNumFmtIds: number[] = [];
    const xfRegex = /<xf\b[^>]*>/g;

    const xfs = cellXfsBlock[0].match(xfRegex) ?? [];
    for (const xf of xfs) {
      const numFmtMatch = xf.match(/numFmtId="(\d+)"/);
      cellXfsNumFmtIds.push(numFmtMatch ? Number(numFmtMatch[1]) : 0);
    }

    return { cellXfsNumFmtIds, numFmts };
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
  // DETECTA FORMATO DE DATA
  // ======================================================
  private isDateFormat(numFmtId: number, formatCode?: string): boolean {
    // Built-ins mais comuns de data
    const builtInDateIds = [14, 15, 16, 17, 22];

    if (builtInDateIds.includes(numFmtId)) return true;

    // Custom: detecta presença de padrões de data/hora
    if (!formatCode) return false;

    const code = formatCode.toLowerCase();

    // remove trechos entre aspas e escapes (pra não confundir)
    const cleaned = code.replace(/"[^"]*"/g, '');

    // Heurística simples: contém d, m, y (e não é apenas texto)
    // Exemplos: dd/mm/yy, m/d/yyyy, dd-mm-yyyy, etc
    const hasDateTokens = /(^|[^a-z])([dmy]){1,}/i.test(cleaned) || /dd|mm|yy|yyyy/.test(cleaned);
    const hasPercent = cleaned.includes('%');

    return hasDateTokens && !hasPercent;
  }

  // ======================================================
  // DETECTA FORMATO DE PERCENTUAL
  // ======================================================
  private isPercentFormat(numFmtId: number, formatCode?: string): boolean {
    // Built-ins percentuais comuns
    const builtInPercentIds = [9, 10];
    if (builtInPercentIds.includes(numFmtId)) return true;

    if (!formatCode) return false;
    return formatCode.includes('%');
  }

  // ======================================================
  // CONVERTE DATA DO EXCEL (1900-based) → dd/MM/yyyy
  // ======================================================
  private excelDateToString(serial: number): string {
    // Excel base (corrige o bug do 1900)
    const excelEpoch = new Date(1899, 11, 30);
    const date = new Date(excelEpoch.getTime() + serial * 86400000);

    const dd = String(date.getDate()).padStart(2, '0');
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const yyyy = date.getFullYear();

    return `${dd}/${mm}/${yyyy}`;
  }
}
