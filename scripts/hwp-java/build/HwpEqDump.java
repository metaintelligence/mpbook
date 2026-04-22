import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPChar;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPCharControlExtend;
import kr.dogfoot.hwplib.object.bodytext.control.Control;
import kr.dogfoot.hwplib.object.bodytext.control.ControlEquation;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.CtrlHeaderGso;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.GsoHeaderProperty;
import kr.dogfoot.hwplib.reader.HWPReader;

public class HwpEqDump {
  public static void main(String[] args) throws Exception {
    HWPFile h = HWPReader.fromFile(args[0]);
    int pidx = 0;
    int eqCount = 0;
    for (Section s : h.getBodyText().getSectionList()) {
      for (Paragraph p : s.getParagraphs()) {
        pidx++;
        if (p.getText()!=null) {
          int i=0;
          for (HWPChar ch : p.getText().getCharList()) {
            if (ch instanceof HWPCharControlExtend) {
              HWPCharControlExtend ce=(HWPCharControlExtend)ch;
              if (ce.isEquation()) {
                System.out.println("PARA "+pidx+" has eq-extend at char "+i+", code="+Integer.toHexString(ch.getCode()));
              }
            }
            i++;
          }
        }
        if (p.getControlList()==null) continue;
        for (Control c : p.getControlList()) {
          if (c.getType()==ControlType.Equation) {
            eqCount++;
            ControlEquation e=(ControlEquation)c;
            CtrlHeaderGso hdr=e.getHeader();
            GsoHeaderProperty gp=hdr.getProperty();
            System.out.println("=== EQ #"+eqCount+" para="+pidx+" ===");
            System.out.println("script="+e.getEQEdit().getScript().toUTF16LEString());
            System.out.println("letterSize="+e.getEQEdit().getLetterSize()+", baseLine="+e.getEQEdit().getBaseLine()+", prop="+e.getEQEdit().getProperty());
            System.out.println("versionInfo="+e.getEQEdit().getVersionInfo().toUTF16LEString()+", font="+e.getEQEdit().getFontName().toUTF16LEString());
            System.out.println("xOff="+hdr.getxOffset()+", yOff="+hdr.getyOffset()+", w="+hdr.getWidth()+", h="+hdr.getHeight()+", z="+hdr.getzOrder());
            System.out.println("likeWord="+gp.isLikeWord()+", applyLineSpace="+gp.isApplyLineSpace()+", allowOverlap="+gp.isAllowOverlap()+", textFlow="+gp.getTextFlowMethod()+", textHorz="+gp.getTextHorzArrange()+", objSort="+gp.getObjectNumberSort());
          }
        }
      }
    }
    System.out.println("TOTAL_EQ="+eqCount);
  }
}
