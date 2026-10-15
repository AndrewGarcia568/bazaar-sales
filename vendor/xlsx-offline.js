/* Minimal offline XLSX writer for Bazaar Sales.
 * Supports the small subset used by this app: json_to_sheet, book_new,
 * book_append_sheet and writeFile. Creates a standards-based XLSX workbook
 * with inline string cells and numeric cells, without network dependencies.
 */
(function (global) {
  'use strict';
  const encoder = new TextEncoder();
  const xmlEscape = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[ch]));
  const colName = n => { let s=''; while(n>0){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);} return s; };
  const bytes = value => encoder.encode(value);
  const concat = chunks => { const total=chunks.reduce((n,c)=>n+c.length,0); const out=new Uint8Array(total); let off=0; for(const c of chunks){out.set(c,off);off+=c.length;} return out; };
  const crc32 = data => { let crc=0xffffffff; for(let i=0;i<data.length;i++){crc^=data[i];for(let j=0;j<8;j++)crc=(crc>>>1)^((crc&1)?0xedb88320:0);} return (crc^0xffffffff)>>>0; };
  const u16 = n => [n&255,(n>>>8)&255];
  const u32 = n => [n&255,(n>>>8)&255,(n>>>16)&255,(n>>>24)&255];
  function zipStore(files) {
    const local=[], central=[]; let offset=0;
    for(const file of files){
      const name=bytes(file.name), data=bytes(file.data), crc=crc32(data);
      const header=concat([new Uint8Array(u32(0x04034b50)),new Uint8Array(u16(20)),new Uint8Array(u16(0)),new Uint8Array(u16(0)),new Uint8Array(u16(0)),new Uint8Array(u16(0)),new Uint8Array(u32(crc)),new Uint8Array(u32(data.length)),new Uint8Array(u32(data.length)),new Uint8Array(u16(name.length)),new Uint8Array(u16(0)),name,data]);
      local.push(header);
      const cdir=concat([new Uint8Array(u32(0x02014b50)),new Uint8Array(u16(20)),new Uint8Array(u16(20)),new Uint8Array(u16(0)),new Uint8Array(u16(0)),new Uint8Array(u16(0)),new Uint8Array(u16(0)),new Uint8Array(u32(crc)),new Uint8Array(u32(data.length)),new Uint8Array(u32(data.length)),new Uint8Array(u16(name.length)),new Uint8Array(u16(0)),new Uint8Array(u16(0)),new Uint8Array(u16(0)),new Uint8Array(u16(0)),new Uint8Array(u32(0)),new Uint8Array(u32(offset)),name]);
      central.push(cdir); offset+=header.length;
    }
    const centralBytes=concat(central), localBytes=concat(local);
    const end=concat([new Uint8Array(u32(0x06054b50)),new Uint8Array(u16(0)),new Uint8Array(u16(0)),new Uint8Array(u16(files.length)),new Uint8Array(u16(files.length)),new Uint8Array(u32(centralBytes.length)),new Uint8Array(u32(localBytes.length)),new Uint8Array(u16(0))]);
    return concat([localBytes,centralBytes,end]);
  }
  function worksheetXml(sheet) {
    const rows=sheet.rows.map((row,ri)=>{
      const cells=row.map((value,ci)=>{
        if(value===null||value===undefined||value==='') return '';
        const ref=colName(ci+1)+(ri+1);
        if(typeof value==='number' && Number.isFinite(value)) return `<c r="${ref}"><v>${value}</v></c>`;
        if(typeof value==='boolean') return `<c r="${ref}" t="b"><v>${value?1:0}</v></c>`;
        return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
      }).join('');
      return `<row r="${ri+1}">${cells}</row>`;
    }).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`;
  }
  function writeFile(workbook, filename) {
    const sheets=workbook.SheetNames.map(name=>({name,sheet:workbook.Sheets[name]}));
    const types=sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('');
    const sheetRefs=sheets.map((s,i)=>`<sheet name="${xmlEscape(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('');
    const rels=sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('');
    const files=[
      {name:'[Content_Types].xml',data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${types}</Types>`},
      {name:'_rels/.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'},
      {name:'xl/workbook.xml',data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheetRefs}</sheets></workbook>`},
      {name:'xl/_rels/workbook.xml.rels',data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`},
      {name:'xl/styles.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'},
      ...sheets.map((s,i)=>({name:`xl/worksheets/sheet${i+1}.xml`,data:worksheetXml(s.sheet)}))
    ];
    const blob=new Blob([zipStore(files)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download=filename; a.style.display='none'; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),1500);
  }
  global.XLSX={utils:{json_to_sheet(data){const headers=data.length?Object.keys(data[0]):[];return {headers,rows:[headers,...data.map(item=>headers.map(key=>item[key]===undefined?'':item[key]))]};},book_new(){return {SheetNames:[],Sheets:{}};},book_append_sheet(book,sheet,name){book.SheetNames.push(name);book.Sheets[name]=sheet;}},writeFile};
})(window);
