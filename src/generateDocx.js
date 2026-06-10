import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign
} from "docx";

const DOT = "............";
const gf = (v, ph = DOT) => (v && String(v).trim()) ? String(v).trim() : ph;

function fmtN(n) { return new Intl.NumberFormat("vi-VN").format(Math.round(n || 0)); }

function numToWords(n) {
  if (!n || n === 0) return "Không đồng";
  n = Math.round(n);
  const u = ["","một","hai","ba","bốn","năm","sáu","bảy","tám","chín"];
  function tri(x) {
    if (!x) return "";
    const h=Math.floor(x/100),t=Math.floor((x%100)/10),o=x%10; let s="";
    if(h) s+=u[h]+" trăm ";
    if(t===0&&o&&h) s+="linh ";
    if(t===1){s+="mười ";if(o===5)s+="lăm ";else if(o)s+=u[o]+" ";}
    else if(t>1){s+=u[t]+" mươi ";if(o===1)s+="mốt ";else if(o===5)s+="lăm ";else if(o)s+=u[o]+" ";}
    else if(o) s+=u[o]+" ";
    return s.trim();
  }
  const ty=Math.floor(n/1e9),tr=Math.floor((n%1e9)/1e6),ng=Math.floor((n%1e6)/1e3),du=n%1e3;
  let r="";
  if(ty) r+=tri(ty)+" tỷ ";
  if(tr) r+=tri(tr)+" triệu ";
  if(ng) r+=tri(ng)+" nghìn ";
  if(du) r+=tri(du);
  r=r.trim();
  return r.charAt(0).toUpperCase()+r.slice(1)+" đồng chẵn";
}

// ─── helpers ───
const BD  = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const NBD = { style: BorderStyle.NONE,   size: 0, color: "FFFFFF" };
const AB  = { top: BD,  bottom: BD,  left: BD,  right: BD  };
const NB  = { top: NBD, bottom: NBD, left: NBD, right: NBD };
const CM  = { top: 80, bottom: 80, left: 100, right: 100 };
const PW  = 10466; // A4 content width (11906 - 2×720 margins) in DXA

function r(text, opts = {})    { return new TextRun({ text: String(text||""), font:"Times New Roman", size:22, ...opts }); }
function rb(text, opts = {})   { return r(text, { bold:true, ...opts }); }
function p(children, opts = {}) {
  return new Paragraph({ children: Array.isArray(children)?children:[children], ...opts });
}
function ep(before=0) { return p([r("")], { spacing:{before,after:0} }); }

function tc(children, w, opts={}) {
  return new TableCell({
    borders: AB, margins: CM,
    width: { size: w, type: WidthType.DXA },
    children: Array.isArray(children)?children:[children],
    ...opts
  });
}

// ─── main export ───
export async function generateQuoteDocx(quote, co) {
  const sub = (quote.items||[]).reduce((s,i)=>s+(i.amount||(i.quantity*i.unitPrice)||0),0)||quote.subtotal||0;
  const vr  = quote.vatRate  ?? 8;
  const vi  = quote.vatIncluded ?? false;
  const va  = vi ? 0 : Math.round(sub*vr/100);
  const tot = quote.total || sub+va;
  const coN = co?.name  || "CÔNG TY CỔ PHẦN QUÀ TẶNG VIVA";
  const coE = co?.email || "lienhe@quatangviva.com";

  // col widths: STT | Tên SP | Hình ảnh | ĐVT | SL | Đơn giá | Thành tiền  (sum = PW)
  const CW = [440, 2960, 1440, 560, 780, 1620, 2666];
  const CW16 = CW.slice(0,6).reduce((a,b)=>a+b,0); // first 6 cols merged width

  const doc = new Document({
    styles: { default: { document: { run: { font:"Times New Roman", size:22 } } } },
    sections: [{
      properties: {
        page: {
          size: { width:11906, height:16838 },
          margin: { top:720, right:720, bottom:720, left:720 }
        }
      },
      children: [

        // ══ HEADER TABLE ══
        new Table({
          width: { size:PW, type:WidthType.DXA },
          columnWidths: [Math.round(PW*0.73), Math.round(PW*0.27)],
          borders: { top:NBD, bottom:{...BD,size:12}, left:NBD, right:NBD, insideH:NBD, insideV:NBD },
          rows: [new TableRow({ children: [
            new TableCell({ width:{size:Math.round(PW*0.73),type:WidthType.DXA}, borders:NB, margins:{top:0,bottom:120,left:0,right:0},
              children:[
                p([rb(coN, {size:26})]),
                p([r("68 Nguyễn Huệ, Phường Sài Gòn, TP Hồ Chí Minh",{size:20})]),
                p([r("VP HCM: 189 Tây Thạnh, Tây Thạnh HCM",{size:20})]),
                p([r("VP HN: 149 Trần Hòa, Định Công Hoàng Mai HN",{size:20})]),
                p([r("Hotline: 1900 8159  |  Email: lienhe@quatangviva.com  |  Website: quatangviva.com",{size:20})]),
              ]
            }),
            new TableCell({ width:{size:Math.round(PW*0.27),type:WidthType.DXA}, borders:NB, verticalAlign:VerticalAlign.CENTER,
              children:[
                p([rb("vivagift",{size:36,color:"2e7d32"})],{alignment:AlignmentType.CENTER}),
                p([r("quatangviva.com",{size:18,color:"555555"})],{alignment:AlignmentType.CENTER}),
              ]
            }),
          ]})]
        }),

        // ══ TITLE ══
        ep(200),
        p([rb("BẢNG BÁO GIÁ SẢN PHẨM & DỊCH VỤ",{size:34})],
          {alignment:AlignmentType.CENTER, spacing:{before:100,after:200}}),

        // ══ CUSTOMER + QUOTE INFO ══
        new Table({
          width:{size:PW,type:WidthType.DXA}, columnWidths:[Math.round(PW*0.55),Math.round(PW*0.45)],
          borders:{top:NBD,bottom:NBD,left:NBD,right:NBD,insideH:NBD,insideV:NBD},
          rows:[new TableRow({children:[
            new TableCell({ width:{size:Math.round(PW*0.55),type:WidthType.DXA}, borders:NB, children:[
              p([rb("Kính gửi (Quotation for): "), rb(gf(quote.customer?.name))]),
              p([rb("Người liên hệ (Attn): "), r(gf(quote.customer?.contact||quote.customer?.phone))]),
              p([rb("Địa chỉ (Address): "), r(gf(quote.customer?.address))]),
              p([rb("Email: "), r(gf(quote.customer?.email))]),
            ]}),
            new TableCell({ width:{size:Math.round(PW*0.45),type:WidthType.DXA}, borders:NB, children:[
              p([rb("Ngày (Date): "), r(gf(quote.date))],{alignment:AlignmentType.RIGHT}),
              p([rb("Số BG (Ref No): "), r(gf(quote.quoteNumber))],{alignment:AlignmentType.RIGHT}),
              p([rb("Nhân viên phụ trách: "), r(coE)],{alignment:AlignmentType.RIGHT}),
            ]}),
          ]})]
        }),

        ep(120),

        // ══ PRODUCT TABLE ══
        new Table({
          width:{size:PW,type:WidthType.DXA}, columnWidths:CW,
          rows:[
            // header row
            new TableRow({ tableHeader:true, children:
              [["STT",null],["Tên SP & Quy cách kỹ thuật",AlignmentType.LEFT],
               ["Hình ảnh\n(Mockup)",null],["ĐVT",null],["Số\nLượng",null],
               ["Đơn giá\n(VNĐ)",null],["Thành tiền\n(VNĐ)",null]]
              .map(([h,al],ci)=> new TableCell({
                borders:AB, margins:CM, width:{size:CW[ci],type:WidthType.DXA},
                shading:{fill:"D9D9D9",type:ShadingType.CLEAR}, verticalAlign:VerticalAlign.CENTER,
                children:[p([rb(h)],{alignment:al||AlignmentType.CENTER})]
              }))
            }),
            // product rows
            ...(quote.items||[]).map((item,idx)=> new TableRow({ children:[
              tc([p([r(String(item.stt||idx+1))],{alignment:AlignmentType.CENTER})], CW[0], {verticalAlign:VerticalAlign.CENTER}),
              tc([p([r(item.name)]), ...(item.code?[p([r("Mã: "+item.code,{size:20,color:"666666"})])]:[])], CW[1]),
              tc([p([r("")],{alignment:AlignmentType.CENTER})], CW[2]),
              tc([p([r(item.unit||"Cái")],{alignment:AlignmentType.CENTER})], CW[3], {verticalAlign:VerticalAlign.CENTER}),
              tc([p([r(fmtN(item.quantity))],{alignment:AlignmentType.CENTER})], CW[4], {verticalAlign:VerticalAlign.CENTER}),
              tc([p([r(fmtN(item.unitPrice))],{alignment:AlignmentType.RIGHT})], CW[5], {verticalAlign:VerticalAlign.CENTER}),
              tc([p([rb(fmtN(item.amount??item.quantity*item.unitPrice))],{alignment:AlignmentType.RIGHT})], CW[6], {verticalAlign:VerticalAlign.CENTER}),
            ]})),
            // subtotal
            new TableRow({children:[
              new TableCell({borders:AB,margins:CM,columnSpan:6,width:{size:CW16,type:WidthType.DXA},
                children:[p([rb("Cộng tiền hàng (Subtotal):")],{alignment:AlignmentType.RIGHT})]}),
              tc([p([rb(fmtN(sub))],{alignment:AlignmentType.RIGHT})],CW[6]),
            ]}),
            // vat
            new TableRow({children:[
              new TableCell({borders:AB,margins:CM,columnSpan:6,width:{size:CW16,type:WidthType.DXA},
                children:[p([r("Thuế GTGT / VAT ("+vr+"%):")],{alignment:AlignmentType.RIGHT})]}),
              tc([p([r(fmtN(va))],{alignment:AlignmentType.RIGHT})],CW[6]),
            ]}),
            // grand total
            new TableRow({children:[
              new TableCell({borders:AB,margins:CM,columnSpan:6,width:{size:CW16,type:WidthType.DXA},
                children:[p([rb("TỔNG CỘNG TIỀN THANH TOÁN (GRAND TOTAL):")],{alignment:AlignmentType.RIGHT})]}),
              tc([p([rb(fmtN(tot))],{alignment:AlignmentType.RIGHT})],CW[6]),
            ]}),
            // amount in words
            new TableRow({children:[
              new TableCell({borders:AB,margins:CM,columnSpan:7,width:{size:PW,type:WidthType.DXA},
                children:[p([r("(Bằng chữ: "),rb(numToWords(tot)),r(")")])]})
            ]}),
          ]
        }),

        // ══ TERMS ══
        ep(220),
        p([rb("ĐIỀU KHOẢN & ĐIỀU KIỆN (TERMS & CONDITIONS)",{size:24})],{spacing:{before:0,after:100}}),
        p([rb("Thanh toán (Payment):")]),
        p([r("• Đợt 1: Thanh toán 50% tổng giá trị báo giá trong vòng 7 ngày khi lên đơn đặt hàng / hai bên ký kết hợp đồng.")]),
        p([r("• Đợt cuối: Thanh toán phần còn lại trong vòng 5 ngày kể từ khi nhận đầy đủ hàng và hóa đơn tài chính hợp lệ.")]),
        p([rb("Thời hạn báo giá (Quotation Valid):"),r(" Báo giá có giá trị trong vòng 15 ngày kể từ ngày báo giá.")]),
        p([rb("Giá trên đã bao gồm:"),r(" Phí in ấn / khắc laze theo yêu cầu, Miễn phí giao hàng đến 1 địa chỉ của Khách hàng.")]),
        p([rb("Yêu cầu file thiết kế:"),r(" Khách hàng cung cấp logo định dạng vector (.AI, .EPS) để đảm bảo chất lượng in ấn / khắc laser sắc nét nhất.")]),
        p([rb("Giao hàng (Delivery):"),r(" Hàng hóa sẽ được giao trong vòng 10-12 ngày làm việc kể từ ngày nhận được tiền thanh toán đợt 1.")]),
        ...(quote.notes?[p([r("📌 "+quote.notes)])]:[]),

        // ══ SIGNATURE ══
        ep(200),
        new Table({
          width:{size:PW,type:WidthType.DXA}, columnWidths:[Math.round(PW*0.55),Math.round(PW*0.45)],
          borders:{top:NBD,bottom:NBD,left:NBD,right:NBD,insideH:NBD,insideV:NBD},
          rows:[new TableRow({children:[
            new TableCell({width:{size:Math.round(PW*0.55),type:WidthType.DXA},borders:NB,children:[ep(0)]}),
            new TableCell({width:{size:Math.round(PW*0.45),type:WidthType.DXA},borders:NB,children:[
              p([rb("BÊN BÁO GIÁ")],{alignment:AlignmentType.CENTER}),
              p([rb(coE)],{alignment:AlignmentType.CENTER}),
              p([r("Nhân viên kinh doanh")],{alignment:AlignmentType.CENTER}),
              ep(1200),
              p([r("(Ký, ghi rõ họ tên)")],{alignment:AlignmentType.CENTER}),
            ]}),
          ]})]
        }),

      ]
    }]
  });

  return await Packer.toBlob(doc);
}
