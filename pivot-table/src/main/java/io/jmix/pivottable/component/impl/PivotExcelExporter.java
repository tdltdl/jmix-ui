/*
 * Copyright 2021 Haulmont.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package io.jmix.pivottable.component.impl;


import io.jmix.core.CoreProperties;
import io.jmix.core.MessageTools;
import io.jmix.core.Messages;
import io.jmix.core.metamodel.datatype.Datatype;
import io.jmix.core.metamodel.datatype.DatatypeRegistry;
import io.jmix.core.metamodel.model.MetaClass;
import io.jmix.core.security.CurrentAuthentication;
import io.jmix.ui.Notifications;
import io.jmix.ui.UiProperties;
import io.jmix.ui.component.ComponentsHelper;
import io.jmix.ui.data.impl.HasMetaClass;
import io.jmix.ui.download.ByteArrayDataProvider;
import io.jmix.ui.download.DownloadFormat;
import io.jmix.ui.download.Downloader;
import io.jmix.ui.screen.ScreenContext;
import io.jmix.uiexport.exporter.excel.ExcelAutoColumnSizer;
import io.jmix.pivottable.component.PivotTable;
import io.jmix.pivottable.model.extension.PivotData;
import io.jmix.pivottable.model.extension.PivotDataCell;
import io.jmix.pivottable.model.extension.PivotDataSeparatedCell;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.config.BeanDefinition;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import javax.annotation.Nullable;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.List;

import static com.google.common.base.Strings.isNullOrEmpty;
import static io.jmix.core.common.util.Preconditions.checkNotNullArgument;


/**
 * Exports {@link PivotData} to XLS file.
 */
@Component("ui_PivotExcelExporter")
@Scope(BeanDefinition.SCOPE_PROTOTYPE)
public class PivotExcelExporter {

    public static final int MAX_ROW_INDEX = 65535;

    /**
     * CAUTION Magic number! This multiplier is used for calculating column width. Without this multiplier all columns
     * in the file will be collapsed.
     */
    protected static final int COLUMN_WIDTH_MULTIPLIER = 48;

    public static final String DEFAULT_FILE_NAME = "pivotData";

    protected HSSFWorkbook wb;
    protected HSSFSheet sheet;

    protected HSSFFont stdFont;

    protected HSSFCellStyle cellLabelBoldStyle;

    protected HSSFCellStyle cellDateTimeStyle;
    protected HSSFCellStyle boldCellDateTimeStyle;

    protected HSSFCellStyle cellDateStyle;
    protected HSSFCellStyle boldCellDateStyle;

    protected HSSFCellStyle cellTimeStyle;
    protected HSSFCellStyle boldCellTimeStyle;

    protected HSSFCellStyle cellIntegerStyle;
    protected HSSFCellStyle boldCellIntegerStyle;

    protected HSSFCellStyle cellDoubleStyle;
    protected HSSFCellStyle boldCellDoubleStyle;

    protected String fileName;
    protected MetaClass entityMetaClass;

    protected Messages messages;
    protected Downloader downloader;

    protected String dateTimeParseFormat;
    protected SimpleDateFormat dateTimeFormatter;
    protected String dateParseFormat;
    protected SimpleDateFormat dateFormatter;
    protected String timeParseFormat;
    protected SimpleDateFormat timeFormatter;

    protected Notifications notifications;

    @Autowired
    protected CoreProperties coreProperties;
    @Autowired
    protected UiProperties uiProperties;
    @Autowired
    protected MessageTools messageTools;
    @Autowired
    protected DatatypeRegistry datatypeRegistry;
    @Autowired
    protected CurrentAuthentication currentAuthentication;

    public PivotExcelExporter(PivotTable pivotTable) {
        init(pivotTable);
    }

    public PivotExcelExporter() {
    }

    public void init(PivotTable pivotTable) {
        entityMetaClass = pivotTable.getDataProvider() instanceof HasMetaClass ?
                ((HasMetaClass) pivotTable.getDataProvider()).getMetaClass() : null;

        initNotifications(pivotTable);
    }

    @Autowired
    public void setMessages(Messages messages) {
        this.messages = messages;
    }

    @Autowired
    public void setDownloader(Downloader downloader) {
        this.downloader = downloader;
    }

    protected void initNotifications(PivotTable pivotTable) {
        ScreenContext screenContext = ComponentsHelper.getScreenContext(pivotTable);
        notifications = screenContext.getNotifications();
    }

    /**
     * Export to Xls.
     *
     * @param pivotData pivot with aggregated data
     * @param fileName  file name
     */
    public void exportPivotTable(PivotData pivotData, String fileName) {
        checkNotNullArgument(pivotData);

        if (isPivotDataEmpty(pivotData)) {
            showNoDataWarning();
            return;
        }

        if (!isNullOrEmpty(fileName)) {
            this.fileName = fileName;
        } else if (entityMetaClass != null) {
            this.fileName = messageTools.getEntityCaption(entityMetaClass);
        } else {
            this.fileName = DEFAULT_FILE_NAME;
        }

        createWorkbookWithSheet();
        createCellsStyle();

        createRows(pivotData);

        if (isXlsMaxRowNumberExceeded(pivotData)) {
            showWarnNotification();
        }

        export(downloader);
    }

    /**
     * Export to Xls.
     *
     * @param pivotData  pivot with aggregated data
     * @param fileName   file name
     * @param downloader ExportDisplay implementation
     */
    public void exportPivotTable(PivotData pivotData, String fileName, Downloader downloader) {
        checkNotNullArgument(pivotData);

        if (isPivotDataEmpty(pivotData)) {
            showNoDataWarning();
            return;
        }

        if (downloader == null) {
            throw new IllegalArgumentException("ExportDisplay is null");
        }

        if (!isNullOrEmpty(fileName)) {
            this.fileName = fileName;
        } else if (entityMetaClass != null) {
            this.fileName = messageTools.getEntityCaption(entityMetaClass);
        } else {
            this.fileName = DEFAULT_FILE_NAME;
        }

        createWorkbookWithSheet();
        createCellsStyle();

        createRows(pivotData);

        if (isXlsMaxRowNumberExceeded(pivotData)) {
            showWarnNotification();
        }

        export(downloader);
    }

    protected void createRows(PivotData pivotData) {
        PivotDataExcelHelper excelUtils = new PivotDataExcelHelper(pivotData);
        List<List<PivotDataSeparatedCell>> dataRows = excelUtils.getRows();

        int columns = excelUtils.getOriginColumnsNumber();
        ExcelAutoColumnSizer[] sizers = columns != -1 ? new ExcelAutoColumnSizer[columns] : null;

        for (int i = 0; i < dataRows.size(); i++) {
            if (i > MAX_ROW_INDEX) {
                break;
            }

            HSSFRow hssfRow = sheet.createRow(i);
            List<PivotDataSeparatedCell> row = dataRows.get(i);
            for (PivotDataSeparatedCell cell : row) {
                HSSFCell hssfCell = hssfRow.createCell(cell.getIndexCol());
                initCell(hssfCell, cell);

                if (sizers != null) {
                    updateColumnSize(sizers, cell);
                }
            }
        }

        if (sizers != null) {
            for (int i = 0; i < sizers.length; i++) {
                sheet.setColumnWidth(i, sizers[i].getWidth() * COLUMN_WIDTH_MULTIPLIER);
            }
        }

        for (CellRangeAddress rangeAddress : excelUtils.getCellRangeAddresses()) {
            sheet.addMergedRegion(rangeAddress);
        }
    }

    protected void initCell(HSSFCell hssfCell, PivotDataSeparatedCell cell) {
        PivotDataCell.Type type = cell.getType();
        switch (type) {
            case DECIMAL:
                BigDecimal bigDecimal = new BigDecimal(cell.getValue());
                Datatype<BigDecimal> bigDecimalDatatype = datatypeRegistry.get(BigDecimal.class);
                String formattedValue = bigDecimalDatatype.format(bigDecimal, currentAuthentication.getLocale());
                try {
                    bigDecimal = bigDecimalDatatype.parse(formattedValue, currentAuthentication.getLocale());
                } catch (ParseException e) {
                    throw new RuntimeException("Unable to parse numeric value", e);
                }
                hssfCell.setCellType(CellType.NUMERIC);
                //noinspection ConstantConditions
                hssfCell.setCellValue(bigDecimal.doubleValue());
                hssfCell.setCellStyle(cell.isBold() ? boldCellDoubleStyle : cellDoubleStyle);
                break;
            case INTEGER:
                Long longValue = Long.parseLong(cell.getValue());
                Datatype<Long> longDatatype = datatypeRegistry.get(Long.class);
                String formattedIntValue = longDatatype.format(longValue, currentAuthentication.getLocale());
                try {
                    longValue = longDatatype.parse(formattedIntValue, currentAuthentication.getLocale());
                } catch (ParseException e) {
                    throw new RuntimeException("Unable to parse numeric value", e);
                }
                hssfCell.setCellType(CellType.NUMERIC);
                //noinspection ConstantConditions
                hssfCell.setCellValue(longValue);
                hssfCell.setCellStyle(cell.isBold() ? boldCellIntegerStyle : cellIntegerStyle);
                break;
            case DATE_TIME:
                initDateTimeCell(hssfCell, cell, dateTimeFormatter, cellDateTimeStyle, boldCellDateTimeStyle);
                break;
            case DATE:
                initDateTimeCell(hssfCell, cell, dateFormatter, cellDateStyle, boldCellDateStyle);
                break;
            case TIME:
                initDateTimeCell(hssfCell, cell, timeFormatter, cellTimeStyle, boldCellTimeStyle);
                break;
            default:
                hssfCell.setCellType(CellType.STRING);
                hssfCell.setCellValue(cell.getValue());
                if (cell.isBold()) {
                    hssfCell.setCellStyle(cellLabelBoldStyle);
                }
                break;
        }
    }

    protected void updateColumnSize(ExcelAutoColumnSizer[] sizers, PivotDataSeparatedCell cell) {
        if (sizers[cell.getIndexCol()] == null) {
            ExcelAutoColumnSizer sizer = new ExcelAutoColumnSizer();
            sizers[cell.getIndexCol()] = sizer;
            sizers[cell.getIndexCol()].notifyCellValue(cell.getValue(), stdFont);
        }

        if (sizers[cell.getIndexCol()].isNotificationRequired(cell.getIndexRow())) {
            sizers[cell.getIndexCol()].notifyCellValue(cell.getValue(), stdFont);
        }
    }

    protected void initDateTimeCell(HSSFCell hssfCell, PivotDataSeparatedCell cell, SimpleDateFormat formatter,
                                    HSSFCellStyle cellStyle, HSSFCellStyle boldCellStyle) {
        if (formatter != null) {
            try {
                hssfCell.setCellValue(formatter.parse(cell.getValue()));
                if (cell.isBold()) {
                    hssfCell.setCellStyle(boldCellStyle);
                } else {
                    hssfCell.setCellStyle(cellStyle);
                }
                return;
            } catch (ParseException e) {
                // ignore because we set it as string
            }
        }
        // set as string
        hssfCell.setCellType(CellType.STRING);
        hssfCell.setCellValue(cell.getValue());
        if (cell.isBold()) {
            hssfCell.setCellStyle(cellLabelBoldStyle);
        }
    }

    protected void createWorkbookWithSheet() {
        wb = new HSSFWorkbook();
        sheet = wb.createSheet("Export");
    }

    protected void createCellsStyle() {
        HSSFFont boldFont = wb.createFont();
        boldFont.setBold(true);

        stdFont = wb.createFont();

        cellLabelBoldStyle = wb.createCellStyle();
        cellLabelBoldStyle.setFont(boldFont);

        String dateTimeFormat = messages.getMessage("pivotExcelExporter.dateTimeFormat");
        cellDateTimeStyle = wb.createCellStyle();
        cellDateTimeStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(dateTimeFormat));
        boldCellDateTimeStyle = wb.createCellStyle();
        boldCellDateTimeStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(dateTimeFormat));
        boldCellDateTimeStyle.setFont(boldFont);

        String dateFormat = messages.getMessage("pivotExcelExporter.timeFormat");
        cellDateStyle = wb.createCellStyle();
        cellDateStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(dateFormat));
        boldCellDateStyle = wb.createCellStyle();
        boldCellDateStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(dateFormat));
        boldCellDateStyle.setFont(boldFont);

        String timeFormat = messages.getMessage("pivotExcelExporter.timeFormat");
        cellTimeStyle = wb.createCellStyle();
        cellTimeStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(timeFormat));
        boldCellTimeStyle = wb.createCellStyle();
        boldCellTimeStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(timeFormat));
        boldCellTimeStyle.setFont(boldFont);

        String integerFormat = messages.getMessage("pivotExcelExporter.integerFormat");
        cellIntegerStyle = wb.createCellStyle();
        cellIntegerStyle.setDataFormat(getBuiltinFormat(integerFormat));
        boldCellIntegerStyle = wb.createCellStyle();
        boldCellIntegerStyle.setDataFormat(getBuiltinFormat(integerFormat));
        boldCellIntegerStyle.setFont(boldFont);

        DataFormat format = wb.createDataFormat();
        String doubleFormat = messages.getMessage("pivotExcelExporter.doubleFormat");
        cellDoubleStyle = wb.createCellStyle();
        cellDoubleStyle.setDataFormat(format.getFormat(doubleFormat));
        boldCellDoubleStyle = wb.createCellStyle();
        boldCellDoubleStyle.setDataFormat(format.getFormat(doubleFormat));
        boldCellDoubleStyle.setFont(boldFont);
    }

    protected short getBuiltinFormat(String format) {
        return HSSFDataFormat.getBuiltinFormat(format);
    }

    protected void showWarnNotification() {
        notifications.create(Notifications.NotificationType.WARNING)
                .withCaption(messages.getMessage("actions.warningExport.title"))
                .withDescription(messages.getMessage("actions.warningExport.message"))
                .withPosition(Notifications.Position.MIDDLE_CENTER)
                .show();
    }

    protected void export(Downloader downloader) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            wb.write(out);
        } catch (IOException e) {
            throw new RuntimeException("Unable to write document", e);
        }
        if (fileName == null) {
            fileName = DEFAULT_FILE_NAME;
        }

        downloader.download(new ByteArrayDataProvider(out.toByteArray(), uiProperties.getSaveExportedByteArrayDataThresholdBytes(),
                        coreProperties.getTempDir()),
                fileName + ".xls", DownloadFormat.XLS);
    }

    protected void showNoDataWarning() {
        notifications.create(Notifications.NotificationType.WARNING)
                .withCaption(messages.getMessage("warningNotification.caption"))
                .withPosition(Notifications.Position.MIDDLE_CENTER)
                .show();
    }

    protected boolean isPivotDataEmpty(PivotData pivotData) {
        return pivotData.getDataNumCols() == 0
                || pivotData.getDataNumRows() == 0;
    }

    /**
     * @param pivotData pivot with aggregated data
     * @return true if exported table contains more than 65536 records
     */
    public boolean isXlsMaxRowNumberExceeded(PivotData pivotData) {
        return MAX_ROW_INDEX < pivotData.getAllRows().size();
    }

    /**
     * @return dateTime format or null
     */
    @Nullable
    public String getDateTimeParseFormat() {
        return dateTimeParseFormat;
    }

    /**
     * Sets dateTime format that will be used to finding dateTime value and exporting it to excel with dateTime type.
     *
     * @param dateTimeParseFormat dateTime format (e.g. dd/MM/yyyy HH:mm)
     */
    public void setDateTimeParseFormat(String dateTimeParseFormat) {
        this.dateTimeParseFormat = dateTimeParseFormat;

        if (!isNullOrEmpty(dateTimeParseFormat)) {
            dateTimeFormatter = new SimpleDateFormat(dateTimeParseFormat);
        }
    }

    /**
     * @return date format or null
     */
    @Nullable
    public String getDateParseFormat() {
        return dateParseFormat;
    }

    /**
     * Sets date format that will be used to finding dateTime value and exporting it to excel with date type. If there
     * is no format set, date properties will be recognized as text value.
     *
     * @param dateParseFormat date format (e.g. dd/MM/yyyy)
     */
    public void setDateParseFormat(String dateParseFormat) {
        this.dateParseFormat = dateParseFormat;

        if (!isNullOrEmpty(dateParseFormat)) {
            dateFormatter = new SimpleDateFormat(dateParseFormat);
        }
    }

    /**
     * @return time format or null
     */
    @Nullable
    public String getTimeParseFormat() {
        return timeParseFormat;
    }

    /**
     * Sets date format that will be used to finding dateTime value and exporting it to excel with date type. If there
     * is no format set, time properties will be recognized as text value.
     *
     * @param timeParseFormat time format (e.g. HH:mm)
     */
    public void setTimeParseFormat(String timeParseFormat) {
        this.timeParseFormat = timeParseFormat;

        if (!isNullOrEmpty(timeParseFormat)) {
            timeFormatter = new SimpleDateFormat(timeParseFormat);
        }
    }
}