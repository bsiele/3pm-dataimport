package threepm;

import java.io.File;

/**
 *
 * @author fegati
 */
public class Main {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        /*JFileChooser fileChooser = new JFileChooser();
        int returnVal = fileChooser.showOpenDialog(null);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            new ExcelReader(
                    fileChooser.getSelectedFile().getAbsoluteFile()).getRowAsListFromExcel();
        }*/
        new ExcelReader(new File("/home/siech/Documents/open_heaven/intelliSOFT/3PM/3pm_V2/targets/Tx_Curr_Targets.xlsx")).getRowAsListFromExcel();
        new ExcelReader(new File("/home/siech/Documents/open_heaven/intelliSOFT/3PM/3pm_V2/targets/Tx_New_Targets.xlsx")).getRowAsListFromExcel();
    }

}
