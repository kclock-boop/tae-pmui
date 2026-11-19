import JSZip from "jszip";
export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export type Field = { id:string; text:string; context:string; suggested:boolean; path:string; index:number };
export type Template = { name:string; size:number; zip:JSZip; xml:Record<string,string>; fields:Field[] };
const elements = (node:Element|Document, name:string) => Array.from(node.getElementsByTagNameNS(W,name));
const text = (node:Element) => Array.from(node.getElementsByTagNameNS(W,"*")).filter(e=>["t","br","cr","tab"].includes(e.localName)).map(e=>e.localName==="t" ? e.textContent||"" : e.localName==="tab" ? "\t" : "\n").join("");
const ancestor = (node:Node, name:string):Element|null => { for(let p=node.parentElement;p;p=p.parentElement) if(p.localName===name && p.namespaceURI===W) return p; return null; };
function parse(xml:string) { const d = new DOMParser().parseFromString(xml,"application/xml"); if (d.getElementsByTagName("parsererror").length) throw new Error("DOCX 내부 XML을 읽을 수 없습니다."); return d; }
export async function openDocx(file:File):Promise<Template> {
  if(!/\.docx$/i.test(file.name)) throw new Error("DOCX 파일만 업로드할 수 있습니다. DOC, HWP, PDF는 Word에서 DOCX로 변환해 주세요.");
  if(file.size>10*1024*1024) throw new Error("10 MB 이하의 DOCX를 선택해 주세요.");
  let zip:JSZip;
  try { zip=await JSZip.loadAsync(await file.arrayBuffer()); } catch { throw new Error("올바른 DOCX 파일이 아닙니다. 암호가 걸린 파일은 암호를 해제해 주세요."); }
  if(!zip.file("word/document.xml") || !zip.file("[Content_Types].xml")) throw new Error("Word 문서 구조가 없는 파일입니다.");
  // ZIP central directory sizes are checked before decompression to bound resource use.
  const entries=Object.values(zip.files); let unpacked=0;
  for(const entry of entries) { const size=(entry as unknown as {_data?:{uncompressedSize?:number}})._data?.uncompressedSize || 0; unpacked+=size; if(size>20*1024*1024 || unpacked>50*1024*1024 || entries.length>2500) throw new Error("압축 해제 용량이 큰 문서입니다. 이미지나 첨부 파일을 줄여 주세요."); }
  const paths=entries.filter(e=>/^word\/(document|header\d+|footer\d+)\.xml$/.test(e.name)).map(e=>e.name);
  const xml:Record<string,string>={}, fields:Field[]=[];
  for(const path of paths) {
    const source=await zip.file(path)!.async("string"); if(/<!DOCTYPE|<!ENTITY/i.test(source)) throw new Error("외부 엔티티가 포함된 문서는 지원하지 않습니다.");
    xml[path]=source; const dom=parse(source); const ps=elements(dom,"p");
    ps.forEach((p,index)=>{
      if(ancestor(p,"p") || ["drawing","object","pict","fldChar","instrText"].some(tag=>elements(p,tag).length)) return;
      const t=text(p), cell=ancestor(p,"tc"), row=ancestor(p,"tr");
      const context=row ? text(row).replace(/\s+/g," ").slice(0,100) || "표의 빈 입력란" : t.slice(0,80) || "빈 문단";
      fields.push({id:`${path}::${index}`,path,index,text:t,context:`${path.includes("header") ? "머리글 · " : path.includes("footer") ? "바닥글 · " : ""}${context}`,suggested:path==="word/document.xml" && (!t.trim() || /\{\{[^}]+\}\}|\[.*?(입력|작성|제목|사유|금액).*?\]/.test(t)) && (!cell || !elements(cell,"vMerge").some(e=>e.getAttributeNS(W,"val")!=="restart"))});
    });
  }
  if(!fields.length) throw new Error("편집 가능한 문단이 없습니다. 이미지로만 된 문서는 지원하지 않습니다.");
  if(fields.length>1200) throw new Error("1,200개 이하 문단의 문서를 사용해 주세요.");
  return {name:file.name,size:file.size,zip,xml,fields};
}
export async function exportDocx(template:Template, values:Record<string,string>):Promise<ArrayBuffer> {
  const zip=(template.zip as JSZip & {clone:()=>JSZip}).clone();
  for(const [path,source] of Object.entries(template.xml)) {
    const changed=template.fields.filter(f=>f.path===path && values[f.id]!==undefined && values[f.id]!==f.text);
    if(!changed.length) continue;
    const dom=parse(source), ps=elements(dom,"p");
    for(const f of changed) {
      const p=ps[f.index];
      const ts=elements(p,"t"); const runs=elements(p,"r");
      // Preserve run properties, paragraph properties, tables and all unrelated package parts.
      for(const e of [...ts,...elements(p,"br"),...elements(p,"cr"),...elements(p,"tab")]) e.remove();
      let run=runs[0]; if(!run) { run=dom.createElementNS(W,"w:r");p.appendChild(run); }
      values[f.id].split("\n").forEach((line,i)=>{
        if(i) run.appendChild(dom.createElementNS(W,"w:br"));
        const t=dom.createElementNS(W,"w:t");t.setAttributeNS("http://www.w3.org/XML/1998/namespace","xml:space","preserve");t.textContent=line;run.appendChild(t);
      });
    }
    zip.file(path,new XMLSerializer().serializeToString(dom));
  }
  return zip.generateAsync({type:"arraybuffer",compression:"DEFLATE"});
}
export async function sampleDocx():Promise<ArrayBuffer> {
  const zip=new JSZip();
  zip.file("[Content_Types].xml",`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("_rels/.rels",`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  const p=(t:string,b=false)=>`<w:p><w:pPr><w:spacing w:after="180"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="맑은 고딕" w:ascii="Arial"/>${b?"<w:b/>":""}<w:sz w:val="22"/></w:rPr><w:t>${t}</w:t></w:r></w:p>`;
  const cell=(t:string,width:number,shade=false)=>`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${shade?'<w:shd w:fill="F1F4F9"/>':''}</w:tcPr>${p(t,shade)}</w:tc>`;
  const rows=[['제목','{{품의 제목}}'],['기안 부서 / 작성자','{{부서 및 작성자}}'],['작성 목적','{{작성 목적}}'],['품의 사유','{{품의 사유}}'],['세부 내용','{{세부 내용}}'],['소요 예산','{{소요 예산}}'],['기대 효과','{{기대 효과}}'],['시행 일정','{{시행 일정}}']];
  zip.file("word/document.xml",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="600"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="40"/></w:rPr><w:t>품 의 서</w:t></w:r></w:p>${p('결재     담당 __________     검토 __________     승인 __________')}<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>${['top','left','bottom','right','insideH','insideV'].map(s=>`<w:${s} w:val="single" w:sz="4" w:color="B6C3D4"/>`).join('')}</w:tblBorders><w:tblCellMar><w:top w:w="150" w:type="dxa"/><w:left w:w="170" w:type="dxa"/><w:bottom w:w="150" w:type="dxa"/><w:right w:w="170" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="7000"/></w:tblGrid>${rows.map(([a,b])=>`<w:tr>${cell(a,2000,true)}${cell(b,7000)}</w:tr>`).join('')}</w:tbl>${p('위와 같이 품의하오니 검토 후 재가하여 주시기 바랍니다.')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1200" w:bottom="1200" w:left="1453" w:right="1453"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({type:"arraybuffer"});
}
