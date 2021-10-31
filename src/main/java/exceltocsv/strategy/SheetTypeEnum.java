package exceltocsv.strategy;

import exceltocsv.strategy.types.Xls;
import exceltocsv.strategy.types.Xlsx;

public enum SheetTypeEnum {
    XLSX("xlsx") {
        @Override
        public SheetType getSheetType() {
            return new Xlsx();
        }
    }, XLS("xls") {
        @Override
        public SheetType getSheetType() {
            return new Xls();
        }
    };

    String fileExtension;

    SheetTypeEnum(String fileExtension) {
        this.fileExtension = fileExtension;
    }

    public abstract SheetType getSheetType();
}
