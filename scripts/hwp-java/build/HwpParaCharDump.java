import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.*;
import kr.dogfoot.hwplib.reader.HWPReader;

public class HwpParaCharDump {
  public static void main(String[] args) throws Exception {
    HWPFile h = HWPReader.fromFile(args[0]);
    int targetPara = Integer.parseInt(args[1]);
    int pidx=0;
    for (Section s: h.getBodyText().getSectionList()) {
      for (Paragraph p: s.getParagraphs()) {
        pidx++;
        if (pidx!=targetPara) continue;
        System.out.println("PARA="+pidx);
        if (p.getText()==null) {System.out.println("no text"); return;}
        int i=0;
        for (HWPChar ch: p.getText().getCharList()) {
          String t=ch.getType().name();
          String extra="";
          if (ch instanceof HWPCharNormal) {
            extra=((HWPCharNormal)ch).getCh()+"";
          } else if (ch instanceof HWPCharControlExtend) {
            HWPCharControlExtend ce=(HWPCharControlExtend)ch;
            extra = "eq="+ce.isEquation()+",gso="+ce.isGSO()+",tbl="+ce.isTable();
          }
          System.out.println(i+"\t"+t+"\tcode="+Integer.toHexString(ch.getCode())+"\t"+extra);
          i++;
        }
        return;
      }
    }
  }
}
