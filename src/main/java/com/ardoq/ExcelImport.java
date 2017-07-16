package com.ardoq;

import com.ardoq.model.*;
import com.ardoq.service.FieldService;
import com.ardoq.util.SyncUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;
import retrofit.RestAdapter;

import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelImport {

    private static final String fieldColMapping_prefix = "fieldColMapping_";
    private static final String compMappingPrefix = "compMapping_";
    private static final String dynamicCompMappingPrefix = "dynamicCompMapping_";
    private static final int CONSECUTIVE_EMPTY_ROWS_MAX = 10;
    private static String componentSeparator = "::";
    private static String token;
    private static String host;
    private static RestAdapter.LogLevel logLevel = RestAdapter.LogLevel.NONE;

    private final static Properties config = new Properties();
    private static Map<String, String> columnMapping = new HashMap<>();
    private static String modelName;
    private static String workspaceName;
    private static String componentSheet;
    private static String componentFile;
    private static String referenceFile;
    private static Map<String, String> compMapping = new HashMap<>();
    private static Map<String, String> dynamicCompMapping = new HashMap<>();


    private static String descriptionColumn;
    private static ArdoqClient client;
    private static SyncUtil ardoqSync;
    private static String organization;
    private static String referenceSheet;
    private static String referenceDefaultLinkType;
    private static int referenceLinkTypeColumn;
    private static int referenceStartFromRow;
    private static int referenceSourceColumn;
    private static int referenceStartFromColumn;

    static HashMap<String, Component> cachedMap = new HashMap<>();

    public static void main(String[] args) throws IOException {
        if (args.length > 0) {
            String configFile = args[0];
            System.out.println("Loading config: " + configFile);
            config.load(new FileReader(configFile));
            parseConfig();
            initClient();
            verifyModel();
            syncComponents();

            if (null != referenceFile) {
                syncReferences();
            }

            if (config.getProperty("deleteMissing", "no").trim().equals("YES")) {
                ardoqSync.deleteNotSyncedItems();
            }

            System.out.println(ardoqSync.getReport());
        } else {
            System.err.println("Path to config file must be first argument.");
        }
    }

    private static void verifyModel() {
        Model model = ardoqSync.getModel();
        String modelId = getModelId(model);
        FieldService fieldService = client.field();
        List<Field> allFields = fieldService.getAllFields();
        List<Field> modelFields = allFields.stream().filter(field -> field.getModel().equals(modelId)).collect(Collectors.toList());
        System.out.println("Fields in Model <" + modelName + "> :");
        toString(modelFields);
        Map<String, String> fieldMapping = new HashMap<>();
        Map<String, String> labelMapping = new HashMap<>();
        for (Field field : modelFields) {
            fieldMapping.put(field.getName(), field.getLabel());
            labelMapping.put(field.getLabel(), field.getName());
        }
        for (String column : columnMapping.keySet()) {
            String field = columnMapping.get(column);
            if (!fieldMapping.keySet().contains(field) && labelMapping.keySet().contains(field)) {
                System.out.println("For column <" + column + "> replacing field <" + field + "> with label <" + labelMapping.get(field) + ">.");
                columnMapping.put(column, labelMapping.get(field));
            }
        }
        List<String> fields = modelFields.stream().map(Field::getName).collect(Collectors.toList());
        List<String> columnFields = new ArrayList<>(columnMapping.values());
        List<String> importingUnknown = new ArrayList<>(columnFields);
        importingUnknown.removeAll(fields);
        List<String> notImporting = new ArrayList<>(fields);
        notImporting.removeAll(columnFields);

        System.out.println("Columns not mapping to fields: " + importingUnknown);
        System.out.println("Fields not mapping to columns: " + notImporting);
    }

    private static String getModelId(Model model) {
        if (model != null) {
            return model.getId();
        } else {
            throw new RuntimeException("Model <" + modelName + "> not found in Ardoq.");
        }
    }

    private static void toString(List<Field> fields) {
        for (Field field : fields) {
            System.out.println("\tName: <" + field.getName() +
                            ">, Label: <" + field.getLabel() +
                            ">, Type: " + field.getType() +
                            ">, Descr: <" + field.getDescription() +
                            ">."
            );
        }
    }

    private static void syncComponents() throws IOException {

        System.out.println("Loading Excel file: " + componentFile);
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(componentFile));

        System.out.println("Finding spread sheet: " + componentSheet);
        // Import components
        XSSFSheet compSheet = workbook.getSheet(componentSheet);

        System.out.println("Analyzing sheet");
        XSSFRow headingRow = compSheet.getRow(0);

        int descriptionIndex = -1;
        HashMap<Integer, String> compTypeMap = new HashMap<>();
        HashMap<Integer, String> fieldTypeMap = new HashMap<>();
        HashMap<Integer, Integer> dynamicCompTypeColumns = new HashMap<>();

        int componentCellRange = 0;
        Iterator<Cell> headingCellIterator = headingRow.cellIterator();
        while (headingCellIterator.hasNext()) {
            Cell headingCell = headingCellIterator.next();
            String heading = headingCell.getStringCellValue().trim();

            if (heading.length() > 0) {
                boolean found = false;
                if (heading.equals(descriptionColumn)) {
                    found = true;
                    descriptionIndex = headingCell.getColumnIndex();
                    System.out.println("Found description column: " + heading + ", " + descriptionIndex);
                }
                if (compMapping.containsKey(heading)) {
                    found = true;
                    System.out.println("Found componentType heading: " + heading + " , " + headingCell.getColumnIndex());
                    compTypeMap.put(headingCell.getColumnIndex(), compMapping.get(heading));
                    componentCellRange = (headingCell.getColumnIndex() > componentCellRange) ? headingCell.getColumnIndex() : componentCellRange;
                    compMapping.remove(heading);
                }
                if (columnMapping.containsKey(heading)) {
                    found = true;
                    System.out.println("Found field heading: " + heading + " , " + headingCell.getColumnIndex());
                    fieldTypeMap.put(headingCell.getColumnIndex(), columnMapping.get(heading));
                    columnMapping.remove(heading);
                }
                if (dynamicCompMapping.containsKey(heading)) {
                    found = true;
                    System.out.println("Found field heading: " + heading + " , " + headingCell.getColumnIndex());
                    componentCellRange = (headingCell.getColumnIndex() > componentCellRange) ? headingCell.getColumnIndex() : componentCellRange;
                    dynamicCompTypeColumns.put(headingCell.getColumnIndex(), findDynamicTypeColumnIndex(headingRow, dynamicCompMapping.get(heading)));
                    columnMapping.remove(heading);
                }
                if (!found) {
                    System.out.println("Ignored column, no component or field mapping found: " + heading);
                }
            }
        }

        for (String key : columnMapping.keySet()) {
            System.out.println("WARNING: Couldn't find Column: " + key + " - cannot map it to field: " + columnMapping.get("key"));
        }

        for (String key : compMapping.keySet()) {
            System.out.println("WARNING: Couldn't find Column: " + key + " - cannot map it to component type: " + compMapping.get("key"));
        }

        int rowIndex = 1;
        int emptyRowCount = 0;
        XSSFRow row = compSheet.getRow(rowIndex);

        while (row != null) {
            ExcelComponent currentComp = null;
            String parentPath = null;
            for (int i = 0; i < componentCellRange + 1; i++) {
                XSSFCell currentCell = row.getCell(i);
                String compType = resolveComponentType(row, compTypeMap, dynamicCompTypeColumns, i);
                if (currentCell != null && compType != null) {
                    String name = (currentCell.getCellType() != Cell.CELL_TYPE_STRING) ?
                            currentCell.getRawValue() : currentCell.getStringCellValue();
                    if (!("".equals(name))) {
                        parentPath = (parentPath == null) ? name : parentPath + SyncUtil.SPLIT_CHARACTER + name;
                        Component c = getComponent(parentPath, name, compType);
                        currentComp = new ExcelComponent(parentPath, c, currentComp);
                    }
                }
            }
            if (currentComp != null) {
                System.out.println("Found component: " + currentComp.getMyComponent().getName());
                currentComp.getMyComponent().setDescription((descriptionIndex != -1 && row.getCell(descriptionIndex) != null) ? row.getCell(descriptionIndex).getStringCellValue() : "");
                HashMap<String, Object> fields = new HashMap<String, Object>();
                for (Integer column : fieldTypeMap.keySet()) {
                    Cell fieldValue = row.getCell(column);
                    if (fieldValue != null) {
                        String key = fieldTypeMap.get(column);
                        fields.put(key, getFieldValue(fieldValue.getCellType(), fieldValue));
                    }
                }
                currentComp.getMyComponent().setFields(fields);
            }
            row = compSheet.getRow(++rowIndex);
            if (row != null) {
                if (row.cellIterator().hasNext()) {
                    emptyRowCount = 0;
                } else {
                    emptyRowCount++;
                }
                if (emptyRowCount == CONSECUTIVE_EMPTY_ROWS_MAX) {
                    row = null;
                }
            }
        }
        for (ExcelComponent ec : ExcelComponent.getRootNodes()) {
            ec.setMyComponent(ardoqSync.addComponent(ec.getMyComponent()));
            //Update cache.
            cachedMap.put(ec.getPath(), ec.getMyComponent());
            storeRecursive(ec);
        }
        System.out.println("DONE syncing components!");
    }

    private static Object getFieldValue(int type, Cell cell) {
        switch (type) {
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case Cell.CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            case Cell.CELL_TYPE_FORMULA:
                return getFieldValue(cell.getCachedFormulaResultType(), cell);
            default:
                return cell.getStringCellValue();
        }
    }

    private static String resolveComponentType(XSSFRow row, HashMap<Integer, String> compTypeMap,
                                               HashMap<Integer, Integer> dynamicCompTypeColumns, int i) {
        if (compTypeMap.containsKey(i)) {
            return compTypeMap.get(i);
        } else if (dynamicCompTypeColumns.containsKey(i)) {
            XSSFCell cell = row.getCell(dynamicCompTypeColumns.get(i));
            if (cell != null && !cell.getStringCellValue().isEmpty()) {
                return cell.getStringCellValue();
            }
            return cell != null ? cell.getStringCellValue() : null;
        }
        return null;
    }

    private static Integer findDynamicTypeColumnIndex(XSSFRow headingRow, String typeColumnHeader) {
        return StreamUtils.asStream(headingRow.iterator())
                .filter(cell -> cell.getCellType() == Cell.CELL_TYPE_STRING &&
                        cell.getStringCellValue().equals(typeColumnHeader))
                .findFirst()
                .orElseThrow(() -> new RuntimeException("Column " + typeColumnHeader + " doesn't exist, unable to map dynamic type!"))
                .getColumnIndex();
    }

    private static void syncReferences() throws IOException {

        System.out.println("Loading Excel file: " + referenceFile);
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(referenceFile));

        System.out.println("Finding reference spread sheet: " + referenceFile);
        // Import components
        XSSFSheet referenceSheet = workbook.getSheet(ExcelImport.referenceSheet);
        System.out.println("Analyzing sheet");

        int rowIndex = referenceStartFromRow;
        XSSFRow referencesRow = referenceSheet.getRow(rowIndex);

        while (referencesRow != null) {
            String sourcePath = getStringValueFromCell(referencesRow, referenceSourceColumn);
            if (sourcePath != null) {
                sourcePath = sourcePath.replace(componentSeparator, SyncUtil.SPLIT_CHARACTER);
                Component sourceComp = cachedMap.get(sourcePath);
                if (null != sourceComp) {
                    List<Component> targetComponents = getTargetComponents(referencesRow);
                    Map<String, Integer> refTypes = ardoqSync.getModel().getReferenceTypes();
                    Integer linkType = refTypes.get(getStringValueFromCell(referencesRow, referenceLinkTypeColumn));
                    if (linkType == null) {
                        linkType = refTypes.get(referenceDefaultLinkType);
                    }
                    if (linkType == null) {
                        linkType = (Integer) refTypes.values().toArray()[0];
                    }

                    for (Component target : targetComponents) {
                        ardoqSync.addReference(new Reference(ardoqSync.getWorkspace().getId(), "", sourceComp.getId(), target.getId(), linkType));
                    }

                } else {
                    System.err.println("Couldn't find source component: " + sourcePath);
                }
            } else {
                System.err.println("Could not find source component in row: " + (rowIndex + 1));
            }
            referencesRow = referenceSheet.getRow(++rowIndex);
        }
        System.out.println("DONE syncing references!");
    }

    private static List<Component> getTargetComponents(XSSFRow referencesRow) {
        ArrayList<Component> comps = new ArrayList<Component>();
        String tc = getStringValueFromCell(referencesRow, referenceStartFromColumn);

        Component c = cachedMap.get(tc.replace(componentSeparator, SyncUtil.SPLIT_CHARACTER));
        if (c != null) {
            comps.add(c);
        } else {
            System.err.println("Couldn't find target component: " + tc);
        }

        return comps;
    }

    private static String getStringValueFromCell(XSSFRow referencesRow, int cellNumber) {
        String cellValue = "";
        Cell sourceCell = referencesRow.getCell(cellNumber);
        if (sourceCell != null) {
            cellValue = sourceCell.getStringCellValue();
        }
        if (cellValue.length() == 0) {
            cellValue = null;
        }
        return cellValue;
    }

    private static void storeRecursive(ExcelComponent ec) {
        for (ExcelComponent child : ec.getChildren()) {
            child.getMyComponent().setParent(ec.getMyComponent().getId());
            child.setMyComponent(ardoqSync.addComponent(child.getMyComponent()));
            //Update cache.
            cachedMap.put(child.getPath(), child.getMyComponent());
            storeRecursive(child);
        }
    }

    private static Component getComponent(String path, String name, String type) {
        Component comp = ardoqSync.getComponentByPath(path);
        if (comp == null) {
            comp = cachedMap.get(path);
        } else if (!comp.getType().equals(type)) {
            comp.setTypeId(ardoqSync.getModel().getComponentTypeByName(type));
            comp.setType(type);
        }

        if (comp == null) {
            comp = new Component(name, ardoqSync.getWorkspace().getId(), "", ardoqSync.getModel().getComponentTypeByName(type));
            cachedMap.put(path, comp);
        } else {
            cachedMap.put(path, comp);
        }
        return comp;
    }

    private static void initClient() {
        System.out.println("Connecting to: " + host + " with token: " + token);
        client = new ArdoqClient(host, token);
        client.setLogLevel(logLevel);
        client.setOrganization(organization);
        List<Workspace> workspaces = client.workspace().findWorkspacesByName(workspaceName);

        if (workspaces.size() > 1) {
            System.out.println("Multiple workspaces match name '" + workspaceName + "'. Please rename or delete workspaces with identical names");
            System.exit(0);
        }

        Workspace workspace = null;
        if (workspaces.size() == 1) {
            workspace = workspaces.get(0);
        }

        if (workspaces.size() == 0) {
            Model template = client.model().getTemplateByName(modelName);
            workspace = new Workspace(workspaceName, template.getId(), "");
            workspace = client.workspace().createWorkspace(workspace);
        }

        ardoqSync = new SyncUtil(client, workspace);
    }

    private static void parseConfig() {
        host = config.getProperty("ardoqHost", "https://app.ardoq.com");
        token = config.getProperty("ardoqToken", System.getenv("ardoqToken"));
        organization = config.getProperty("organization", "ardoq");
        if (config.getProperty("clientLogLevel") != null) {
            logLevel = RestAdapter.LogLevel.valueOf(config.getProperty("clientLogLevel"));
        }

        componentSeparator = config.getProperty("referenceComponentSeparator", componentSeparator);

        if (token == null) {
            System.err.println("No ardoqToken specified in property-file or in environment variable ardoqToken.");
            System.exit(-1);
        }

        modelName = getRequiredValue("modelName");

        workspaceName = getRequiredValue("workspaceName");
        componentSheet = getRequiredValue("componentSheet");

        componentFile = getRequiredValue("componentFile");
        descriptionColumn = getRequiredValue("compDescriptionColumn");

        referenceFile = config.getProperty("referenceFile", null);
        referenceSheet = config.getProperty("referenceSheet", null);
        referenceDefaultLinkType = config.getProperty("referenceDefaultLinkType", null);
        referenceLinkTypeColumn = getNumberConfig("referenceLinkTypeColumn");
        referenceStartFromRow = getNumberConfig("referenceStartFromRow");
        referenceSourceColumn = getNumberConfig("referenceSourceColumn");
        referenceStartFromColumn = getNumberConfig("referenceStartFromColumn");

        for (Object o : config.keySet()) {
            String key = (String) o;
            if (key.startsWith(fieldColMapping_prefix)) {
                System.out.println("Mapping column <" + key.replace(fieldColMapping_prefix, "") + "> to field <" + config.getProperty(key) + ">.");
                columnMapping.put(key.replace(fieldColMapping_prefix, ""), config.getProperty(key));
            }

            if (key.startsWith(compMappingPrefix)) {
                System.out.println("Mapping column <" + key.replace(compMappingPrefix, "") + "< to component type <" + config.getProperty(key) + ">.");
                compMapping.put(key.replace(compMappingPrefix, ""), config.getProperty(key));
            }

            if (key.startsWith(dynamicCompMappingPrefix)) {
                System.out.println("Mapping column <" + key.replace(dynamicCompMappingPrefix, "") + ">,  will use dynamic type mapping");
                dynamicCompMapping.put(config.getProperty(key), key.replace(dynamicCompMappingPrefix, ""));
            }
        }
    }

    private static int getNumberConfig(String numericPropertyKey) {
        String s = config.getProperty(numericPropertyKey, "-1").trim();
        Integer i = -1;
        try {
            i = Integer.parseInt(s);
        } catch (NumberFormatException nfe) {
            System.err.println("WARNING! Couldn't parse " + numericPropertyKey + " as Integer. Value was: " + s);
        }

        return i;
    }

    private static String getRequiredValue(String key) {
        String value = config.getProperty(key);
        if (value == null) {
            System.err.println(key + " was not configured in config. Exiting!");
            System.exit(-1);
        }
        return value;
    }
}
