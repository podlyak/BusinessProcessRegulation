//  [ПМ-094] Регламент бизнем процесса
//  Версия: "1.11.05"
//  Дата изменения: 10.11.2024 19:57

void execute() {
    new BusinessProcessRegulationScript(context: context).execute()
}

import groovy.util.logging.Slf4j
import org.apache.commons.collections4.CollectionUtils
import org.apache.commons.lang3.StringUtils
import org.apache.poi.openxml4j.util.ZipSecureFile
import org.apache.poi.util.Units
import org.apache.poi.xwpf.usermodel.*
import org.apache.xmlbeans.XmlCursor
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*
import ru.nextconsulting.bpm.descriptor.annotations.Description
import ru.nextconsulting.bpm.dto.FullModelDefinition
import ru.nextconsulting.bpm.dto.NodeId
import ru.nextconsulting.bpm.dto.SimpleMultipartFile
import ru.nextconsulting.bpm.dto.request.EdgeDefinitionSearch
import ru.nextconsulting.bpm.dto.request.EraseRequest
import ru.nextconsulting.bpm.dto.search.SearchRequest
import ru.nextconsulting.bpm.dto.search.SearchVisibility
import ru.nextconsulting.bpm.repository.*
import ru.nextconsulting.bpm.repository.business.AttributeValue
import ru.nextconsulting.bpm.repository.structure.*
import ru.nextconsulting.bpm.script.repository.TreeRepository
import ru.nextconsulting.bpm.script.tree.elements.Edge
import ru.nextconsulting.bpm.script.tree.elements.ObjectElement
import ru.nextconsulting.bpm.script.tree.node.Model
import ru.nextconsulting.bpm.script.tree.node.ObjectDefinition
import ru.nextconsulting.bpm.script.tree.node.TreeNode
import ru.nextconsulting.bpm.scriptengine.context.ContextParameters
import ru.nextconsulting.bpm.scriptengine.context.CustomScriptContext
import ru.nextconsulting.bpm.scriptengine.exception.SilaScriptException
import ru.nextconsulting.bpm.scriptengine.script.GroovyScript
import ru.nextconsulting.bpm.scriptengine.serverapi.*
import ru.nextconsulting.bpm.scriptengine.util.ParamUtils
import ru.nextconsulting.bpm.scriptengine.util.SilaScriptParameter
import ru.nextconsulting.bpm.scriptengine.util.SilaScriptParameters
import ru.nextconsulting.bpm.utils.JsonConverter

import javax.imageio.ImageIO
import java.awt.image.BufferedImage
import java.lang.reflect.Constructor
import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths
import java.text.SimpleDateFormat
import java.util.concurrent.atomic.AtomicInteger
import java.util.regex.Pattern
import java.util.stream.Collectors
import java.util.stream.Stream
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

/**
 * Скрипт: {@value #SCRIPT_NAME}
 * Скрипт: {@value #SCRIPT_VERSION}
 */
@SilaScriptParameters([
        @SilaScriptParameter(
                name = 'Глубина детализации регламента',
                type = SilaScriptParamType.SELECT_STRING, selectStringValues = ['3 уровень', '4 уровень'],
                defaultValue = '3 уровень'
        ),
])
@Slf4j
class BusinessProcessRegulationScript implements GroovyScript {
    static void main(String[] args) {
        ContextParameters parameters = ContextParameters.builder()
                .login('superadmin')
                .password('WM_Sila_123')
                .apiBaseUrl('http://localhost:8080/')
                .build()
        CustomScriptContext context = CustomScriptContext.create(parameters)

        ScriptParameter modelParam = ScriptParameter.builder()
                .paramType(SilaScriptParamType.NODE)
                .name('modelId')
                .value(JsonConverter.writeValueAsJson(NodeId.builder()
                        .repositoryId('bcc41132-013f-45ce-8e73-f0a095f51ca5')
                        .id('1a8132f0-a43b-11e7-05b7-db7cafd96ef7')
                        .build())
                )
                .build()
        ScriptParameter elementsIdsParam = ScriptParameter.builder()
                .paramType(SilaScriptParamType.STRING_LIST)
                .name('elementsIdsList')
                .value('["1a82b990-a43b-11e7-05b7-db7cafd96ef7", "1a829281-a43b-11e7-05b7-db7cafd96ef7"]')
                .build()
        context.getParameters().add(modelParam)
        context.getParameters().add(elementsIdsParam)

        BusinessProcessRegulationScript script = new BusinessProcessRegulationScript(context: context)
        script.execute()
    }

    CustomScriptContext context
    public static final String SCRIPT_NAME = "[ПМ-094] Регламент бизнес-процесса"
    public static final String SCRIPT_VERSION = "1.11"
    public static final String SCRIPT_DATE = "21.10.2024 15:26"
    private static final String FIRST_LEVEL_ID = "1a8132f0-a43b-11e7-05b7-db7cafd96ef7"

    //private static final String POSITION = "Должность"
    private static final String POSITION_ID = "OT_POS"                  // Должность
    // private static final String FUNCTION = "Функция"
    private static final String FUNCTION_ID = "OT_FUNC"                 // Функция
    private static final String FUNCTION_SYMBOL_ID = "ST_FUNC"
    private static final String GROUP_ID = "OT_GRP"// Группа
    private static final String INFO_CARR_ID = "OT_INFO_CARR"           // Носитель информации
    private static final String OT_PERS_TYPE = "OT_PERS_TYPE"           // Бизнес - роль
    private static final String OT_EVT = "OT_EVT"                       // Событие
    private static final String ROLE_ID = "objectType.personType"
    private static final String OT_ORG_UNIT = "OT_ORG_UNIT"             // Организационная единица
    private static final String OT_APPL_SYS_TYPE = "OT_APPL_SYS_TYPE"   // Тип прикладной системы
    private static final String STATUS_SYMBOL_ID = "d6e8a7b0-7ce6-11e2-3463-e4115bf4fdb9"
    private static final String EPC_ID = "MT_EEPC"
    private static final String PSD_ID = "MT_PRCS_SLCT_DIA"
    private static final String ST_VAL_ADD_CHN_SML_2 = "ST_VAL_ADD_CHN_SML_2"   //Цепь создания добавленной стоимости
    private static final String ST_GRP = "ST_GRP" // Тип символа группа
    private static final String ST_EV = "ST_EV" // Тип симовола интерфейс процесса
    private static final String ST_PRCS_IF = "ST_PRCS_IF" // Тип
    private static final String ST_SCENARIO = "ST_SCENARIO" // Тип символа Сценарий
    private static final String ST_FUNC = "ST_FUNC" //Тип символа функция
    private static final String ST_SOLAR_FUNC = "ST_SOLAR_FUNC" // Тип символа функция SAP
    private static final String ST_PROCESS_TYPE = "22e410f0-e498-11e2-69e4-ac8112d1b401" //Тип символа процесс типовой
    private static final String ST_PROCESS_SAP_TYPE = "9ad2efb0-c1a4-11e4-3864-ff0f8fe73e88" //Тип символа процесс типовой
    private static final String ST_PROCESS_SAP = "f46e6b10-c1a1-11e4-3864-ff0f8fe73e88" // тип символа процесс sap
    private static final String GROUP_I_ST = "fd841c20-cc37-11e6-05b7-db7cafd96ef7" // группировка интерфейсов
    private static final String REL_PROCESS = "75d9e6f0-4d1a-11e3-58a3-928422d47a25"
    private static final String REL_I_PROCESS = "75f2e570-bdd3-11e5-05b7-db7cafd96ef7"//интерфейс смежного процесса
    private static final String ST_GROUP_1_ID = "9a651320-a385-11e3-3864-ff0f8fe73e88"
    private static final String ST_GROUP_2_ID = "87ac9820-a385-11e3-3864-ff0f8fe73e88"
    private static final String ST_PRCS_1 = "ST_PRCS_1" // Тип символа процесс
    private static List<String> stWithout = Arrays.asList(GROUP_I_ST.toUpperCase(), REL_PROCESS.toUpperCase(), REL_I_PROCESS.toUpperCase(), ST_GRP.toUpperCase(),
            ST_EV.toUpperCase(), ST_PRCS_IF.toUpperCase(), ST_GROUP_1_ID.toUpperCase(), ST_GROUP_2_ID.toUpperCase(),
            ST_PROCESS_SAP.toUpperCase(), ST_PROCESS_TYPE.toUpperCase(), ST_PROCESS_SAP_TYPE.toUpperCase())
    // перечисление символ id функций которые нужно исключить


    //  Константы для выбора папки
    private static final String PRIMARY_FOLDER_PATH = "Модели в разработке"
    private static final String SECONDARY_FOLDER_PATH = "Актуальные модели"
    private static final String TEMPLATE_FOLDER_ID = "file-folder-root-id"
    private static Map<String, FullModelDefinition> cacheFullModelDefinition = new HashMap<>()
    private static Map<String, ObjectDefinitionNode> cacheObjectDefinition = new HashMap<>()
    private static Map<String, List<ObjectElementNode>> linkElementsEntriesCache = new HashMap<>()
    private static Map<String, ObjectElementNode> sourceCache = new HashMap<>()
    private static Map<String, ObjectElementNode> targetCache = new HashMap<>()
    private static Map<String, BigInteger> bulletNumberingMap = [:]
    private static Map<String, String> linkTypeNameCache = new HashMap<>()
    public static final Comparator<ObjectElementNode> CHILD_COMPARATOR = Comparator.comparing({ y -> ((ObjectElementNode) y).getY() })
            .thenComparing(Comparator.comparing({ x -> ((ObjectElementNode) x).getX() }))
            .thenComparing(Comparator.comparing({ y -> ((ObjectElementNode) y).getName() }))
    public static final Comparator<ObjectElementNode> ROOT_COMPARATOR = Comparator.comparing({ s -> (((ObjectElementNode) s).getGroup() != null) ? ((ObjectElementNode) s).getGroup().getY() : 0.0d })
            .thenComparing(Comparator.comparing({ rootX -> ((ObjectElementNode) rootX).getRoot().getX() }))
            .thenComparing(Comparator.comparing({ rootY -> ((ObjectElementNode) rootY).getRoot().getY() }))
            .thenComparing(Comparator.comparing({ x -> ((ObjectElementNode) x).getX() }))
            .thenComparing(Comparator.comparing({ y -> ((ObjectElementNode) y).getY() }))
            .thenComparing(Comparator.comparing({ y -> ((ObjectElementNode) y).getName() }))
    public static final Comparator<ObjectElementNode> PSD_COMPARATOR = Comparator.comparing { y -> trimName(((ObjectElementNode) y).getName()) }

    public static String conjunction = "И"
    public static int detailLevel = 3
    public static final Pattern onlyNumberPattern = Pattern.compile('^[\\d/]*\\d$')
    public static final OrgUnitFlatComparator unitFlatComparator = new OrgUnitFlatComparator()
    public static final OrgUnitTreeComparator unitTreeComparator = new OrgUnitTreeComparator()
    public static final Comparator<ObjectDefinitionNode> objectDefinitionNodeComparator = Comparator.comparing(o -> getNodeFullName(o as ObjectDefinitionNode))

    private static Map<String, List<EdgeDefinitionNode>> relationsTargetMap
    private static Map<String, ObjectDefinitionNode> objectMap

    private static final int CM1 = 567 // 1 сантиметр (отступ в документе)
    private static final int CM05 = (int) (CM1 / 2) // 0,5 сантиметра (отступ в документе)
    private static final String bulletSymbol = '•' // • или 
    private static final String bulletFont = "Times New Roman" // Times New Roman или Symbol
    private static final String BLT = bulletSymbol + ' '
    private static BigInteger documentNumbering

    public static boolean FAST_ORG_STRUCT = false
    public static boolean CREATE_RELATIONS = false
    public static boolean DEBUG_LOCAL = false
    public static String LOCAL_PATH = ""

    @Override
    void execute() {
        if (cacheFullModelDefinition.size() > 0) {
            cacheFullModelDefinition.clear()
        }
        if (cacheObjectDefinition.size() > 0) {
            cacheObjectDefinition.clear()
        }
        if (linkElementsEntriesCache.size() > 0) {
            linkElementsEntriesCache.clear()
        }
        if (linkTypeNameCache.size() > 0) {
            linkTypeNameCache.clear()
        }
        if (sourceCache.size() > 0) {
            sourceCache.clear()
        }
        if (targetCache.size() > 0) {
            targetCache.clear()
        }
        if (DEBUG_LOCAL) {
            loadCacheFromFile()
        }

        try {
            log.info("Запуск скрипта $SCRIPT_NAME $SCRIPT_VERSION от $SCRIPT_DATE")

            String deep = ParamUtils.parse(context.findParameter('Глубина детализации регламента')) as String
            String templateName = "reg_bp.docx"
            boolean isFirstLevel = context.modelId().id == FIRST_LEVEL_ID

            TreeRepository repository = context.createTreeRepository(true)

            TreeNode repositoryTreeNode = repository.read(context.modelId().repositoryId, context.modelId().repositoryId)
            TreeNode primaryFolder = findFolder(repository, repositoryTreeNode.id, repositoryTreeNode.id, PRIMARY_FOLDER_PATH)
            TreeNode folder = primaryFolder == null ? findFolder(repository, repositoryTreeNode.id, repositoryTreeNode.id, SECONDARY_FOLDER_PATH) : primaryFolder

            /**
             * Признак,что при построении дерева орг. структуры будем использовать определения связей
             * Если признак есть, то нам загрузить орг. структуру в кэш
             */
            if (FAST_ORG_STRUCT) {
                TreeNode orgStructFolder = findFolder(repository, repositoryTreeNode.id, repositoryTreeNode.id, "Организационная структура")
                loadOrgStruct(orgStructFolder)
            }

            HashMap<String, FileNodeDTO> files = new HashMap<>()
            for (int i = 0; i < context.nodesIdsList().size(); i++) {
                String fileName = "ReglBP_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + "_" + (i + 1)
                String diagramElementId = context.elementsIdsList().get(i)
                files.put(fileName, createBusinessProcessDoc(folder, deep, templateName, diagramElementId, isFirstLevel, fileName, i))
            }

            String fileName = ("ReglBP_" +
                    new SimpleDateFormat("yyyyMMdd HHmmss").format(new Date())).replace(" ", "_")
            FileNodeDTO result = null
            String format = "docx"

            if (files.size() > 1) {
                byte[] zipFile = createZipWithDocs(files)
                format = "zip"

                def fileRepositoryId = 'file-folder-root-id'
                def userId = context.principalId

                def fileNode = FileNodeDTO.builder()
                        .nodeId(NodeId.builder().id(UUID.randomUUID().toString()).repositoryId(fileRepositoryId).build())
                        .parentNodeId(NodeId.builder().id(String.valueOf(userId)).repositoryId(fileRepositoryId).build())
                        .extension(format)
                        .file(new SimpleMultipartFile(fileName, zipFile))
                        .name(fileName + "." + format)
                        .build()
                result = fileNode
            } else if (files.size() == 1) {
                result = files.values().stream().findFirst().get()
            }

            if (result != null) {
                if (DEBUG_LOCAL) {
                    context.setResultFile(result.file.bytes, format, fileName)
                } else {
                    context.getApi(FileApi).uploadFile(result)
                    context.setResultFile(result.file.bytes, format, fileName)
                }
            }

        } finally {
            if (DEBUG_LOCAL) {
                saveCacheToFile()
            }
        }

        println()
    }

    /**
     * Кэшируем орг. структуру
     * @param orgStructFolder папка "Организационная структура"
     */
    void loadOrgStruct(TreeNode orgStructFolder) {
        TreeApi treeApi = context.getApi(TreeApi.class)

        // Получить орг. структуру в виде плоского списка объектов
        def flatTree = treeApi.getAllChild(orgStructFolder.repositoryId, orgStructFolder.id)
        def orgStructFlatTree = flatTree.groupBy { it.type }

        def relationsList = orgStructFlatTree.get(NodeType.EDGE).stream()
                .map(EdgeDefinitionNode.class::cast).toList() as List<EdgeDefinitionNode>

        relationsTargetMap = relationsList.groupBy { it.getTargetObjectDefinitionId() }

        List<ObjectDefinitionNode> objectList = orgStructFlatTree.get(NodeType.OBJECT).stream()
                .map(ObjectDefinitionNode.class::cast).toList() as List<ObjectDefinitionNode>

        objectMap = objectList.collectEntries { [it.getNodeId().getId(), it] }
    }


    static byte[] createZipWithDocs(HashMap<String, FileNodeDTO> files) {
        ByteArrayOutputStream bos = new ByteArrayOutputStream()
        ZipOutputStream zipOut = new ZipOutputStream(bos)

        files.each { fileName, content ->
            ZipEntry zipEntry = new ZipEntry(fileName + ".docx")
            zipOut.putNextEntry(zipEntry)
            zipOut.write(content.file.bytes, 0, content.file.bytes.length)
            zipOut.closeEntry()
        }

        zipOut.close()
        bos.close()

        return bos.toByteArray()
    }

    FileNodeDTO createBusinessProcessDoc(TreeNode folder, String deep, String templateName, String diagramElementId, boolean isFirstLevel, String fileName, int nodeIdIndex) {
        try {
            ModelApi modelApi = context.getApi(ModelApi.class)
            EdgeTypeApi edgeTypeApi = context.getApi(EdgeTypeApi.class)
            ObjectTypeApi objectTypeApi = context.getApi(ObjectTypeApi.class)
            ObjectsApi objectsApi = context.getApi(ObjectsApi.class)
            TreeApi treeApi = context.getApi(TreeApi.class)
            RepositoryApi repositoryApi = context.getApi(RepositoryApi.class)

            String repositoryId = context.modelId().repositoryId
            String nodeId = context.nodesIdsList().get(nodeIdIndex).id

            Node node = treeApi.getNode(repositoryId, nodeId, "true")
            TreeRepository repository = context.createTreeRepository(false)
            Model model = repository.read(context.modelId())
            FullModelDefinition firstLevelModel = getModelDefinition(modelApi, context.modelId().repositoryId, FIRST_LEVEL_ID)
            FullModelDefinition fullModel = getModelDefinition(modelApi, context.modelId().repositoryId, context.modelId().id)
            RepositoryNode repositoryNode = repositoryApi.byId(context.nodeId().getRepositoryId())
            String presetId = repositoryNode.getPresetId()

            if (fullModel == null) {
                throw new SilaScriptException("Функция не найдена")
            }

            detailLevel = Integer.parseInt(deep.replaceAll("[^0-9]+", ""))

            Map<String, OrgUnitNode> visited = new HashMap<String, OrgUnitNode>()
            Map<String, OrgUnitNode> rootNodes = new HashMap<>()

            ObjectElement el = model.getElement(diagramElementId)
            ObjectDefinition object = el.getObjectDefinition()
            ObjectDefinitionNode objectNode = getObjectDefinition(objectsApi, repositoryId, object.id)
            String processName = getNodeFullName(objectNode)

            if (!objectNode.getObjectTypeId().equalsIgnoreCase(FUNCTION_ID)) {
                throw new SilaScriptException("Объект для запуска скрипта не является функцией")
            }
            ObjectElementNode elementNode = new ObjectElementNode(fullModel, objectNode, null, diagramElementId, edgeTypeApi, presetId)

            Set<String> orgUnitNames = linkElementsEntries(elementNode, OT_ORG_UNIT, null, "Выполняет", presetId, null).stream()
                    .map { it.getName() }
                    .collect(Collectors.toSet())


            // Ищем связанные объекты типа "Функция" с типом связи "Подчиняет по процессу"
            //List<String> functionNames = linkElementsName(el, FUNCTION_ID, "Подчиняет по процессу", presetId, edgeTypeApi)

            NodeDefinitionInstanceInfo instanceInfo = fullModel.getModel().getParentNodesInfo().stream().filter { it.getNodeAllocations().stream().anyMatch { location -> location.getModelId().equalsIgnoreCase(FIRST_LEVEL_ID) } }.findFirst().orElse(null)
            ObjectElementNode upperNode = isFirstLevel ? null : getObjectElementsNode(firstLevelModel, getObjectDefinition(objectsApi, repositoryId, instanceInfo.getNodeId()), null, edgeTypeApi, presetId).stream().findFirst().orElse(null)
            // Имя верхн.процесса
            String upperProcessName = isFirstLevel ? processName : instanceInfo.getNodeName()

            List<String> businessOwner = isFirstLevel ? new ArrayList<String>() : linkElementsEntries(upperNode, OT_ORG_UNIT, null, "Выполняет", presetId, null).stream()
                    .map { it.getName() }
                    .collect(Collectors.toList())
            businessOwner.sort()
            businessOwner.stream().forEachOrdered { orgUnitNames::add }
            businessOwner.addAll(orgUnitNames)
            businessOwner.sort()
            businessOwner.replaceAll { BLT + it }


            //получаем декомпозицию в зависимости от уровня 8:16
            List<FullModelDefinition> subDecomposition = getSubDecomposition(fullModel, object, folder, isFirstLevel)
            List<FullModelDefinition> subProcessModel = isFirstLevel ? subDecomposition : Arrays.asList(fullModel)

            Set<ObjectElementNode> sortedGroups = new TreeSet<>((o1, o2) -> o1.getName() <=> o2.getName())
            Map<String, FullModelDefinition> decSubDecomposition = getDecSubDecomposition(sortedGroups, subProcessModel, objectNode, elementNode, presetId, repositoryId, isFirstLevel)

            // == TODO распутать этот фрагмент кода
            Map<String, ObjectElementNode> functions = new HashMap<>()
            Map<String, ObjectElementNode> functionsAll = new HashMap<>()
            Map<String, ObjectDefinitionNode> nearProcessNode = new HashMap<>()

            subProcessModel.each {
                def fns = findFunctionByLevelDecomposition(null, functionsAll, nearProcessNode, it, presetId, detailLevel,
                        repositoryId, repository, isFirstLevel, objectNode.getNodeId().id)
                functions.putAll(fns)
            }

            // Смежный процесс
            List<String> nearProcess = getNearProcess(nearProcessNode, firstLevelModel, detailLevel, isFirstLevel, repositoryId)

            Map<String, ObjectElementNode> subProcessors = functions.values().collectEntries { [it.root.getObjectDefinitionNode().nodeId.id, it.root] }

            Set<ObjectElementNode> subProcessRoles = functions.values().stream()
                    .map { s -> s.getParent() }
                    .filter { parent -> parent != null }
                    .collect(Collectors.toSet())

            functionsAll.values().forEach {
                decSubDecomposition.put(it.getModelDefinition().getModel().nodeId.id, it.getModelDefinition())
            }
            // ==


            Map<String, List<String>> businessRolesForFunctions = new HashMap<>()

            long start1 = System.currentTimeMillis()
            Map<String, Map<String, Object>> preliminaryFunctionData = getPreliminaryFunctionData(functions, presetId, visited, rootNodes, businessRolesForFunctions, detailLevel)
            log.info("Дерево построено за {} мс", System.currentTimeMillis() - start1)
//                rootNodes.values().stream().filter {!it.isPosition}.forEach {it.printTree()}

            // #77206 Создать определения связей по отношениям объектов на моделях
            if (!FAST_ORG_STRUCT && CREATE_RELATIONS) {
                new RelationCreator().createRelationsInForest(rootNodes)
            }

            List<ObjectElementNode> sortedObjects = new ArrayList<>()
            List<String> subProcessName = getSubProcessName(subProcessors, functions, isFirstLevel, sortedObjects)

            Map<String, Integer> functionNumberMap = getFunctionNumberMap(subProcessRoles, subProcessName)

            Map<String, Map<String, Object>> additionalFunctionData = getAdditionalFunctionData(preliminaryFunctionData, functionNumberMap, presetId)

            Map<String, String> parentProcessForFunction = getParentProcessForFunction(functions)

            String p21 = isFirstLevel ? "Процесс «$processName» является процессом верхнего уровня"
                    : "Процесс «$processName» является подпроцессом процесса «$upperProcessName»"
            String p3level4 = "todelete"
            String p3Table = "todelete"
            String p31level4Head = "todelete"
            String p31level4 = "todelete"
            String p31Table = "todelete"
            String verDoc = "todelete"
            String verOt = "todelete"
            String table22 = ""
            String table23 = ""

            List<List<String>> table3 = null
            List<List<String>> table4 = null

            Map<Integer, List<List<String>>> numberTablesToData = new HashMap<>()
            if (detailLevel == 4) {
                p3level4 = "Перечень структурных подразделений и закрепленных за ними ролей процесса «$processName» приведены в таблице 1."
                p3Table = "\\BТаблица 1. Перечень структурных подразделений, должностей и закрепленных за ними ролей."
                p31level4Head = "3.1. Функциональные роли участников процесса"
                p31level4 = "Перечень ролей и ответственность участников процесса «$processName» приведены в Таблице 2."
                p31Table = "\\BТаблица 2. Функциональные роли и ответственность участников процесса «$processName»"

                // п.3 УЧАСТНИКИ ПРОЦЕССА, Таблица 1
                table3 = fillOrganizationalUnitsTable(rootNodes)

                if (isFirstLevel) {
                    table22 = "\\BНаименование подпроцесса"
                    table23 = "\\BНаименование процедуры"
                } else {
                    table22 = "\\BНаименование процедуры"
                    table23 = "\\BНаименование операции"
                }
                table4 = fillFunctionalRolesTable(businessRolesForFunctions, parentProcessForFunction, functions, isFirstLevel)
            }

            numberTablesToData.put(3, table3)
            numberTablesToData.put(4, table4)

            Map<String, String> parameters = new HashMap<>()
            parameters.put("Имя верх.процесса", upperProcessName)
            parameters.put("УПРАВЛЕНИЕ ОРГАНИЗАЦИЕЙ И НОРМИРОВАНИЕМ ТРУДА", processName)
            parameters.put("Управление организацией и нормированием труда", processName)
            parameters.put("ИМЯ ПРОЦЕССА", processName)
            parameters.put("Имя процесса", processName)
            parameters.put("Пункт 3 уровень 4", p3level4)
            parameters.put("Таблица 1", p3Table)
            parameters.put("Пункт 3.1 уровень 4 заголовок", p31level4Head)
            parameters.put("Пункт 3.1 уровень 4", p31level4)
            parameters.put("Таблица 2", p31Table)
            parameters.put("Пункт 2.1", p21)
            parameters.put("версия_документа", verDoc)
            parameters.put("версия_от", verOt)
            parameters.put("Таблица 2 столбец 2", table22)
            parameters.put("Таблица 2 столбец 3", table23)


            // TODO написать что это
            Map<String, List<List<String>>> createTablesToData = fillNewFunctionTable(additionalFunctionData, subProcessRoles, subProcessName)
            decSubDecomposition.entrySet().removeIf(entry ->
                    entry.getValue().getModel().nodeId.id == model.id ||
                            subDecomposition.stream()
                                    .anyMatch(subProcess -> subProcess.getModel().nodeId.id == entry.getValue().getModel().nodeId.id)
            )

            log.info("Получение диаграммы контекста запуска")
            List<ProcessPng> subProcessesPng = processPng(subDecomposition, sortedObjects)
            log.info("Получение диаграмм декомпозиций")
            List<ProcessPng> decSubProcessesPng = processPng(decSubDecomposition.values(), sortedObjects)


            // Состав документа
            int appDocumentCount = subProcessesPng.size() + decSubProcessesPng.size() + 1
            List<String> docContent = getDocContent(functions, appDocumentCount, presetId, repositoryId)


            int appGroupCount = appDocumentCount + 1

            // Подразделение
            List<String> subdivisionsList = getOrganizationalUnitsAndInParameterLists(rootNodes, visited, sortedGroups, appGroupCount, presetId)

            def orgUnitModel = linkElements(elementNode, OT_ORG_UNIT, null, presetId, null, null)
            def gi = getGroupInfo(rootNodes, sortedGroups, orgUnitModel, appDocumentCount, modelApi, repositoryId)
            List<String> groupInfo = gi.get(0) // Состав групп
            List<String> groupAndOrgUnit = gi.get(1) // Подразделение и группы

            // Орг единицы для листа согласования
            List acceptList = orgUnitNames.stream().map { it + " ПАО «Газпром»" }.toList()


            Map<String, List<String>> parameterLists = new HashMap<>()
            // Если связанных процессов нет, параметр не отображается TODO почему??
            parameterLists.put("Смежный процесс", nearProcess)
            parameterLists.put("Подразделение", subdivisionsList)
            parameterLists.put("Владелец бизнес-процесса", addSemicolon(businessOwner, 1))
            parameterLists.put("Функции", functions.values().stream().map { it.getName() }.collect(Collectors.toList()))
            parameterLists.put("Имя подпроцесса", subProcessName)

            parameterLists.put("состав_документов", addSemicolon(docContent, 1))
            parameterLists.put("состав_групп", addSemicolon(groupInfo, 1))
            parameterLists.put("Подразделение и группы", addSemicolon(groupAndOrgUnit, 2))

            parameters.put("Номер_состав_документов", String.valueOf(appDocumentCount))
            parameters.put("Номер_состав_групп", String.valueOf(appGroupCount))

            // Удаление лишних пробелов из значений в parameters
            parameters.replaceAll((k, v) -> trimName(v))

            // Удаление лишних пробелов из значений в parameterLists
            parameterLists.forEach((k, v) -> v.replaceAll(this::trimName))
            //parameterLists = parameterLists.collectEntries{key, value -> [key, value.collect{trimName(it)}]}

            DocumentData documentData = new DocumentData()
            documentData.setParameters(parameters)
            documentData.setCreateTablesToData(createTablesToData)
            documentData.setParameterLists(parameterLists)
            documentData.setNumberTablesToData(numberTablesToData)
            documentData.setSubProcesses(subProcessesPng)
            documentData.setDecSubProcesses(decSubProcessesPng)
            documentData.setOrgUnitNames(acceptList)

            DocumentCreator documentCreator = new DocumentCreator()
            return documentCreator.createDoc(templateName, fileName, documentData)
        } catch (Exception e) {
            log.error("Ошибка создания регламента", e)
            if (e instanceof SilaScriptException) {
                throw e
            }
        }
    }

    static List<String> addSemicolon(List<String> list, int level) {
        for (int i = 0; i < list.size(); i++) {
            String cur = list.get(i)
            if (StringUtils.countMatches(cur, BLT) == level) {
                if (i == list.size() - 1) {
                    list.set(i, cur + '.')
                } else {
                    String next = list.get(i + 1)
                    if (StringUtils.countMatches(next, BLT) == level) {
                        list.set(i, cur + ';')
                    } else {
                        list.set(i, cur + '.')
                    }
                }
            }
        }
        return list
    }

    private Map<String, FullModelDefinition> getDecSubDecomposition(Set<ObjectElementNode> sortedGroups,
                                                                    final List<FullModelDefinition> subProcessModel,
                                                                    ObjectDefinitionNode objectNode,
                                                                    ObjectElementNode elementNode,
                                                                    String presetId,
                                                                    String repositoryId,
                                                                    boolean isFirstLevel) {
        ModelApi modelApi = context.getApi(ModelApi.class)
        EdgeTypeApi edgeTypeApi = context.getApi(EdgeTypeApi.class)
        ObjectTypeApi objectTypeApi = context.getApi(ObjectTypeApi.class)

        Map<String, FullModelDefinition> decSubDecomposition = new HashMap<>()
        subProcessModel.forEach {
            if (isFirstLevel) {
                sortedGroups.addAll(findGroup(it, edgeTypeApi, presetId))
            }

            findDecomposition(it, null, presetId, edgeTypeApi, objectTypeApi, decSubDecomposition, modelApi, repositoryId, isFirstLevel, objectNode.getNodeId().id, detailLevel)
        }

        if (!isFirstLevel) {
            List<ObjectElementNode> grp = new ArrayList<>()
            grp.addAll(linkElementsEntries(elementNode, GROUP_ID, null, null, presetId, null))
            grp.removeAll(Collections.singleton(null))

            sortedGroups.addAll(grp)
        }

        decSubDecomposition.values().each {
            decSubModel -> sortedGroups.addAll(findGroup(decSubModel, edgeTypeApi, presetId))
        }

        return decSubDecomposition
    }

    private List<FullModelDefinition> getSubDecomposition(FullModelDefinition fullModel, ObjectDefinition object, TreeNode folder, boolean isFirstLevel) {
        ModelApi modelApi = context.getApi(ModelApi.class)

        List<FullModelDefinition> subDecomposition = fullModel.getObjects()
                .find { it -> it.getNodeId().id == object._getNodeId().id }
                .getModelAssignments().stream().map { assigment ->
            if (assigment.getModelId() != fullModel.getModel().getNodeId().id) {
                if (!isFirstLevel || assigment.getModelPath().contains(folder.getName())) {
                    getModelDefinition(modelApi, object._getNodeId().repositoryId, assigment.getModelId())
                }
            }
        }.collect(Collectors.toList())
        subDecomposition.removeAll(Collections.singleton(null))
        return subDecomposition
    }

    private List<String> getSubProcessName(Map<String, ObjectElementNode> subProcessors, Map<String, ObjectElementNode> functions, boolean isFirstLevel, List<ObjectElementNode> sortedObjects) {
        List<String> subProcessName = new ArrayList<>()

        List<ObjectElementNode> sortSubProcess = isFirstLevel ?
                subProcessors.values().toList() :
                subProcessors.values().stream().findFirst().orElse(null)?.children?.stream()?.collect(Collectors.toList())

        if (sortSubProcess != null) {
            Collections.sort(sortSubProcess, ROOT_COMPARATOR)
            sortSubProcess.stream().forEach {
                setChildrenNameList(it, subProcessName, sortedObjects, 1, functions.keySet())
            }
        }

        return subProcessName
    }

    private Map<String, Integer> getFunctionNumberMap(Set<ObjectElementNode> subProcessRoles, List<String> subProcessName) {
        Map<String, Integer> functionNumberMap = new HashMap<>()

        Comparator<ObjectElementNode> comparator = Comparator.comparing({ ObjectElementNode subProcNode ->
            {
                String subProc = subProcessName.find { sub -> sub.containsIgnoreCase(subProcNode.getName()) }
                return subProc == null ? -1 : subProcessName.indexOf(subProc)
            }
        })

        subProcessRoles.forEach(subProcessRole -> {
            if (subProcessRole != null) {
                int i = 1
                List<ObjectElementNode> children = subProcessRole.children.stream().collect(Collectors.toList())
                Collections.sort(children, CHILD_COMPARATOR)

                // Обрабатываем каждую функцию
                for (ObjectElementNode function : children) {
                    functionNumberMap.put(function.getId(), i++)
                }
            }
        })
        return functionNumberMap
    }

    private Map<String, String> getParentProcessForFunction(Map<String, ObjectElementNode> functions) {
        Map<String, String> parentProcessForFunction = new HashMap<>()

        for (String functionId : functions.keySet()) {
            ObjectElementNode functionElement = functions.get(functionId)

            // Найти родительский объект для функции TODO
            ObjectElementNode parentProcesses = functionElement.getParent()

            if (parentProcesses != null) {
                // Предполагаем, что у каждой функции может быть только один родительский процесс
                String parentProcessName = parentProcesses.getName()
                parentProcessForFunction.put(functionId, parentProcessName)
            }
        }

        return parentProcessForFunction
    }


    private Map<String, Map<String, Object>> getAdditionalFunctionData(Map<String, Map<String, Object>> preliminaryFunctionData, Map<String, Integer> functionNumberMap, String presetId) {
        EdgeTypeApi edgeTypeApi = context.getApi(EdgeTypeApi.class)
        Map<String, Map<String, Object>> additionalFunctionData = new HashMap<>(preliminaryFunctionData)
        // Новая структура данных
        for (Map<String, Object> functionData : preliminaryFunctionData.values()) {
            String functionId = (String) functionData.get("functionId")
            List<ObjectElementNode> results = (List<ObjectElementNode>) functionData.get("result")
            Map<String, List<ObjectElementNode>> functionForEvents = (Map<String, List<ObjectElementNode>>) functionData.get("functionForEvent")
            String resultConjunction = (String) functionData.get("resultConjunction")
            if (resultConjunction == null || resultConjunction.isEmpty()) {
                resultConjunction = conjunction
            }


            // Создать список результатов с указанием пункта
            List<String> resultList = new LinkedList<>()
            if (results != null && !results.isEmpty()) {
                results.eachWithIndex { event, index ->
                    String punkt = null
                    ObjectElementNode rule = linkElements(event, "OT_RULE", "оценивается с помощью", presetId, "outgoing", edgeTypeApi).stream().findFirst().orElse(null)
                    if (rule != null) {
                        List<ObjectElementNode> functionsEvent = linkElements(rule, FUNCTION_ID, "активизирует", presetId, "outgoing", edgeTypeApi)
                        if (functionsEvent != null && !functionsEvent.isEmpty()) {
                            functionsEvent.forEach { functionEvent ->
                                punkt = punkt == null ? functionNumberMap.get(functionEvent.getId()) : punkt + functionNumberMap.get(functionEvent.getId())
                            }
                        }
                    } else {
                        ObjectElementNode functionEvent = linkElements(event, FUNCTION_ID, "активизирует", presetId, "outgoing", edgeTypeApi).stream().findFirst().orElse(null)
                        if (functionEvent != null) {
                            punkt = functionNumberMap.get(functionEvent.getId())
                        }
                    }
                    String resultText = event.name + (punkt != null ? " (Переход к п. " + punkt + ")" : "")
                    resultList.add(resultText)

                    if (functionForEvents != null && !functionForEvents.isEmpty()) {
                        List<ObjectElementNode> functionForEvent = functionForEvents.get(event.getId())
                        if (functionForEvent != null && !functionForEvent.isEmpty()) {

                            functionForEvent.forEach { functionForEventName ->
                                if (functionForEventName != null) {
                                    String text = "Переход к процессу: " + functionForEventName.getName()
                                    resultList.add(text)
                                }
                            }
                        }
                    }

                    // Добавление союза между результатами, если необходимо
                    if (index < results.size() - 1) {
                        resultList.add(resultConjunction)
                    }
                }
            }

            // Сохранить обновленные данные в структуру additionalFunctionData
            Map<String, Object> existingFunctionData = additionalFunctionData.getOrDefault(functionId, new HashMap<>())
            existingFunctionData.put("results", resultList)
            additionalFunctionData.put(functionId, existingFunctionData)
        }

        return additionalFunctionData
    }

    private Map<String, Map<String, Object>> getPreliminaryFunctionData(final Map<String, ObjectElementNode> functions,
                                                                        String presetId,
                                                                        Map<String, OrgUnitNode> visited,
                                                                        Map<String, OrgUnitNode> rootNodes,
                                                                        Map<String, List<String>> businessRolesForFunctions,
                                                                        int detailLevel
    ) {
        EdgeTypeApi edgeTypeApi = context.getApi(EdgeTypeApi.class)
        ObjectsApi objectsApi = context.getApi(ObjectsApi.class)
        Map<String, Map<String, Object>> preliminaryFunctionData = new HashMap<>()

        for (String functionId : functions.keySet()) {
            ObjectElementNode functionNode = functions.get(functionId)

            List<ObjectElementNode> businessRoles = linkElementsEntries(functionNode, OT_PERS_TYPE, null, null, presetId, null)

            List<String> businessRoleNames = new ArrayList<>()
            if (detailLevel == 4) {
                if (businessRoles != null && !businessRoles.isEmpty()) {
                    for (ObjectElementNode businessRole : businessRoles) {
                        businessRoleNames.add(businessRole.getName())

                        // Создание списка организационных единиц для каждой бизнес-роли
                        List<ObjectElementNode> positions = linkElementsEntries(businessRole, POSITION_ID, null, null, presetId, null)

                        if (positions != null && !positions.isEmpty()) {
                            for (ObjectElementNode position : positions) {
                                String positionId = position.getObjectDefinitionNode().nodeId.id

                                OrgUnitNode positionNode
                                if (visited.containsKey(positionId)) {
                                    // Если узел уже был посещен, используем его
                                    positionNode = visited.get(positionId)
                                } else {
                                    // Иначе создаем новый узел
                                    positionNode = new OrgUnitNode(position.getObjectDefinitionNode(), true)
                                    visited.put(positionNode.id, positionNode)
                                }

                                positionNode.getBusinessRoles().add(getNodeFullName(businessRole.getObjectDefinitionNode()))
                                addPositionToTree(rootNodes, positionNode, objectsApi)
                            }
                        }
                    }
                }
            } else {
                businessRoles = new ArrayList<>()
                businessRoles.addAll(linkElementsEntries(functionNode, OT_ORG_UNIT, null, null, presetId, null))
                if (businessRoles != null && !businessRoles.isEmpty()) {
                    for (ObjectElementNode businessRole : businessRoles) {
                        String orgUnitId = businessRole.getObjectDefinitionNode().nodeId.id

                        OrgUnitNode orgUnitNode
                        if (visited.containsKey(orgUnitId)) {
                            // Если узел уже был посещен, используем его
                            orgUnitNode = visited.get(orgUnitId)
                        } else {
                            // Иначе создаем новый узел
                            orgUnitNode = new OrgUnitNode(businessRole.getObjectDefinitionNode(), false)
                            visited.put(orgUnitNode.id, orgUnitNode)
                        }
                        addPositionToTree(rootNodes, orgUnitNode, objectsApi)
                    }
                }
                businessRoles.addAll(linkElementsEntries(functionNode, GROUP_ID, null, null, presetId, null))
                businessRoles.forEach {
                    businessRoleNames.add(it.getName())
                }
            }


            if (!businessRoleNames.isEmpty()) {
                businessRolesForFunctions.put(functionId, businessRoleNames)
            }

            // Дополнительные данные для новой таблицы
            Map<String, Object> functionData = new HashMap<>()

            // TODO если статусы пустые не выводить слово статус
            // Входящие документы и их статусы
            List<ObjectElementNode> incomingDocs = linkElements(functionNode, INFO_CARR_ID, null, presetId, "incoming", edgeTypeApi)
            Map<String, String> incomingDocsWithStatus = incomingDocs.collectEntries { doc -> [doc.getName(), getStatus(doc, edgeTypeApi, presetId)] }
            functionData.put("incomingDocsWithStatus", incomingDocsWithStatus)

            // Исходящие документы и их статусы
            List<ObjectElementNode> outgoingDocs = linkElements(functionNode, INFO_CARR_ID, null, presetId, "outgoing", edgeTypeApi)
            Map<String, String> outgoingDocsWithStatus = outgoingDocs.collectEntries { doc -> [doc.getName(), getStatus(doc, edgeTypeApi, presetId)] }
            functionData.put("outgoingDocsWithStatus", outgoingDocsWithStatus)

            String resultConjunction = "И"
            // Получение результатов (событий)
            List<ObjectElementNode> results = new ArrayList<>()
            results.addAll(linkElements(functionNode, OT_EVT, null, presetId, "outgoing", edgeTypeApi))

            // Получение правил для определения союза "И" или "ИЛИ"
            List<ObjectElementNode> rules = getAllRules(functionNode, "outgoing", presetId, edgeTypeApi)
            if (rules != null && !rules.isEmpty()) {
                rules.forEach { it ->
                    results.addAll(linkElements(it, OT_EVT, null, presetId, "outgoing", edgeTypeApi))
                }
                // Проверка типа правила (здесь предполагается, что есть метод для получения типа правила)
                if (rules.any { it.getName().containsIgnoreCase("ИЛИ") }) {
                    if (rules.any { it.getName().containsIgnoreCase("И") }) {
                        resultConjunction = "ИЛИ/И"
                    } else {
                        resultConjunction = "ИЛИ"
                    }
                }
            }

            Map<String, List<ObjectElementNode>> functionForEvent = new HashMap<>()
            results.forEach { event ->
                functionForEvent.put(event.getId(), new ArrayList<ObjectElementNode>())
                List<ObjectElementNode> rulesForEvent = getAllRules(event, "outgoing", presetId, edgeTypeApi)
                if (!rulesForEvent.isEmpty()) {
                    rulesForEvent.forEach { it ->
                        List<ObjectElementNode> allFuncEvents = linkElements(it, FUNCTION_ID, null, presetId, "outgoing", edgeTypeApi)
                        if (allFuncEvents != null) {
                            functionForEvent.get(event.getId()).addAll(allFuncEvents.stream().filter { it.getObjectInstance().getSymbolId() == ST_PRCS_IF }.collect(Collectors.toList()))
                        }
                    }
                }

                List<ObjectElementNode> funcForEvent = linkElements(event, FUNCTION_ID, null, presetId, "outgoing", edgeTypeApi)
                if (funcForEvent != null) {
                    functionForEvent.get(event.getId()).addAll(funcForEvent.stream().filter { it.getObjectInstance().getSymbolId() == ST_PRCS_IF }.collect(Collectors.toList()))
                }

                if (resultConjunction == "ИЛИ" && rulesForEvent.any { it.getName().containsIgnoreCase("И") }) {
                    resultConjunction = "ИЛИ/И"
                }
            }

            /*List<String> resultNames = results.stream()
                    .map(result -> result.getName())
                    .collect(Collectors.toList())*/
            functionData.put("result", results)
            functionData.put("resultConjunction", resultConjunction)

            /*List<String> functionForEventNames = results.stream()
                    .map(result -> result.getName())
                    .collect(Collectors.toList())*/
            functionData.put("functionForEvent", functionForEvent)

            // Информационная система
            List<ObjectElementNode> systemsList = linkElements(functionNode, OT_APPL_SYS_TYPE, null, presetId, null, edgeTypeApi)
            List<String> systemList = systemsList.stream()
                    .map(system -> system.getName())
                    .collect(Collectors.toList())
            functionData.put("systems", systemList)

            List<String> responsible = new ArrayList<>(businessRoleNames)
            /*if (detailLevel == 3) {
                responsible = linkElements(functionNode, OT_ORG_UNIT, null, presetId, null, edgeTypeApi)
                        .stream().map { it.getName() }.collect(Collectors.toList())
            }*/

            functionData.put("responsible", responsible)
            functionData.put("functionName", functionNode.getName())
            functionData.put("functionId", functionId)

            // Сохранение дополнительных данных для функции
            preliminaryFunctionData.put(functionId, functionData)
        }

        return preliminaryFunctionData
    }

    private List<ObjectElementNode> getAllRules(ObjectElementNode node, String connectionType, String presetId, EdgeTypeApi edgeTypeApi) {
        List<ObjectElementNode> rules = new ArrayList<>()
        List<ObjectElementNode> rulesForEvent = linkElements(node, "OT_RULE", null, presetId, connectionType, edgeTypeApi)
        if (rulesForEvent != null && !rulesForEvent.isEmpty()) {
            rules.addAll(rulesForEvent)
            rulesForEvent.forEach { it ->
                rules.addAll(getAllRules(it, connectionType, presetId, edgeTypeApi))
            }
        }

        return rules
    }

    private List<ProcessPng> processPng(Collection<FullModelDefinition> subDecomposition, List<ObjectElementNode> sortedObjects) {
        List<ProcessPng> subProcessesPng = new ArrayList<>()
        ModelApi modelApi = context.getApi(ModelApi.class)

        // Нужно очень хитро отсортировать функции.
        // Но такую сортировку мы уже делали для п. 4. И там мы заготовили уже отсортированные объекты - sortedObjects
        // Поэтому задача сводится к следующему:
        // 1. Сопоставить модель объекту
        // 2. Использовать готовый отсортированный список объектов для сортировки моделей

        /// === Начало сортировки
        // Готовим отображение id объекта -> объект - для удобства поиска
        Map<String, ObjectElementNode> sortedObjectsMap = sortedObjects.collectEntries{[it.id, it]}

        // 1. Сопоставить модель объекту
        HashMap<FullModelDefinition, ObjectElementNode> map = new HashMap<>()
        for (FullModelDefinition model : subDecomposition) {
            // ищем для определения модели объекты, которые на эту модель ссылаются
            ObjectElementNode node = model.getModel().getParentNodesInfo().stream()
                    .filter { sortedObjectsMap.get(it.nodeId) != null }
                    .map { it -> sortedObjectsMap.get(it.nodeId) }
                    .find() as ObjectElementNode

            map.put(model, node)
        }

        // 2. Использовать готовый отсортированный список объектов для сортировки моделей
        Map<FullModelDefinition, ObjectElementNode> sorted = map.sort { o1, o2 -> {
            int idx1 = o1.value != null ? sortedObjects.findIndexOf {it.id == o1.value.id} : Integer.MAX_VALUE
            int idx2 = o2.value != null ? sortedObjects.findIndexOf {it.id == o2.value.id} : Integer.MAX_VALUE
            return Integer.compare(idx1, idx2)
        }
        }
        /// === Конец сортировки

        for (FullModelDefinition sub : sorted.keySet()) {
            byte[] png = new byte[0]
            try {
                byte[] pngSub = modelApi.getModelPng(sub.model.getNodeId().repositoryId, sub.model.getNodeId().id)
                if (pngSub != null && pngSub.size() > 0) {
                    png = pngSub
                }
            } catch (Exception e) {
                log.error("Ошибка получения схемы модели", e)
            } finally {
                ProcessPng subProcess = new ProcessPng(sub.getModel().getName(), png)
                subProcessesPng.add(subProcess)
            }
        }

        return subProcessesPng
    }

    private List<String> getNearProcess(Map<String, ObjectDefinitionNode> nearProcessNode, FullModelDefinition firstModel, int detailLevel, boolean isFirstLevel, String repositoryId) {
        ModelApi modelApi = context.getApi(ModelApi.class)
        Set<String> func = new LinkedHashSet<>() // #76593 Убрать дубли в пункте 2.3 Смежные процессы
        Map<String, ObjectDefinitionNode> uppObjectModelAssigmentsName = new HashMap<>()
        firstModel.objects.forEach { object ->
            object.modelAssignments.forEach {
                uppObjectModelAssigmentsName.put(it.getModelId(), object)
            }
        }

        Map<String, String> upperName = new HashMap<>()
        Map<String, String> subName = new HashMap<>()
        Map<String, List<String>> uppSubId = new HashMap<>()
        /*def entries = linkElementsEntries(elementNode, FUNCTION_ID, null, "Подчиняет по процессу", presetId, null)
                .stream()
                .map { BLT + it.getName().replaceAll("\\s+", " ") }
                .collect(Collectors.toSet())
        func.addAll(entries)*/
        nearProcessNode.values().forEach { near ->
            // Если в экземплярах объекта, найденного в п. 1 есть "Проект модели процессов верхнего уровня (справочная модель)",
            // то указать в качестве смежного процесса полное имя объекта. Иначе переходим к п.3.
            if (near.getObjectEntries() != null && near.getObjectEntries().modelId.contains(firstModel.getModel().getNodeId().id)) {
                upperName.put(near.getNodeId().id, getNodeFullName(near))
                if (!uppSubId.containsKey(near.getNodeId().id)) {
                    uppSubId.put(near.getNodeId().id, new ArrayList<String>())
                }
                return
            }

            if (detailLevel == 3) {
                List<ParentModelOfObjectDefinition> objectEntries = near.getObjectEntries().stream()
                        .filter { uppObjectModelAssigmentsName.containsKey(it.getModelId()) }.collect(Collectors.toList())
                if (objectEntries != null && !objectEntries.isEmpty() && objectEntries.stream()
                        .anyMatch { modelOfObject ->
                            modelOfObject.getObjectInstanceInfoList().stream()
                                    .anyMatch { it.getSymbolId().equalsIgnoreCase(ST_VAL_ADD_CHN_SML_2) }
                        }) {
                    objectEntries.stream()
                            .filter { modelOfObject ->
                                modelOfObject.getObjectInstanceInfoList().stream()
                                        .anyMatch { it.getSymbolId().equalsIgnoreCase(ST_VAL_ADD_CHN_SML_2) }
                            }
                            .forEach {
                                subName.put(near.getNodeId().id, getNodeFullName(near))
                                ObjectDefinitionNode upperNode = uppObjectModelAssigmentsName.get(it.getModelId())
                                upperName.put(upperNode.getNodeId().getId(), getNodeFullName(upperNode))
                                if (!uppSubId.containsKey(upperNode.getNodeId().id)) {
                                    uppSubId.put(upperNode.getNodeId().id, new ArrayList<String>())
                                }
                                uppSubId.get(upperNode.getNodeId().id).add(near.getNodeId().id)
                            }

                } else {
                    objectEntries = near.getObjectEntries().stream()
                            .filter { it.getModelTypeId().equalsIgnoreCase(PSD_ID) }
                            .collect(Collectors.toList())
                    if (objectEntries != null && !objectEntries.isEmpty() && objectEntries.stream()
                            .anyMatch { modelOfObject ->
                                modelOfObject.getObjectInstanceInfoList().stream()
                                        .anyMatch { it.getSymbolId().equalsIgnoreCase(ST_SCENARIO) }
                            }) {
                        objectEntries.stream()
                                .filter { modelOfObject ->
                                    modelOfObject.getObjectInstanceInfoList().stream()
                                            .anyMatch { it.getSymbolId().equalsIgnoreCase(ST_SCENARIO) }
                                }
                                .forEach {
                                    FullModelDefinition parent = getModelDefinition(modelApi, repositoryId, it.getModelId())
                                    parent.getModel().getParentNodesInfo().forEach { instanceInfo ->
                                        instanceInfo.getNodeAllocations().forEach {
                                            if (uppObjectModelAssigmentsName.containsKey(it.getModelId())) {
                                                subName.put(instanceInfo.getNodeId(), instanceInfo.getNodeName())
                                                ObjectDefinitionNode upperNode = uppObjectModelAssigmentsName.get(it.getModelId())
                                                upperName.put(upperNode.getNodeId().getId(), getNodeFullName(upperNode))
                                                if (!uppSubId.containsKey(upperNode.getNodeId().id)) {
                                                    uppSubId.put(upperNode.getNodeId().id, new ArrayList<String>())
                                                }
                                                uppSubId.get(upperNode.getNodeId().id).add(instanceInfo.nodeId)
                                            }
                                        }
                                    }
                                }
                    }
                }
            }
            else {
                List<ParentModelOfObjectDefinition> objectEntries = near.getObjectEntries().stream()
                        .filter { it.getModelTypeId().equalsIgnoreCase(PSD_ID) }.collect(Collectors.toList())
                if (objectEntries != null && !objectEntries.isEmpty() && objectEntries.stream()
                        .anyMatch { modelOfObject ->
                            modelOfObject.getObjectInstanceInfoList().stream()
                                    .anyMatch { it.getSymbolId().equalsIgnoreCase(ST_PRCS_1) }
                        }) {
                    objectEntries.stream()
                            .filter { modelOfObject ->
                                modelOfObject.getObjectInstanceInfoList().stream()
                                        .anyMatch { it.getSymbolId().equalsIgnoreCase(ST_PRCS_1) }
                            }
                            .forEach {
                                FullModelDefinition parent = getModelDefinition(modelApi, repositoryId, it.getModelId())
                                parent.getModel().getParentNodesInfo().forEach { instanceInfo ->
                                    instanceInfo.getNodeAllocations().forEach {
                                        if (uppObjectModelAssigmentsName.containsKey(it.getModelId())) {
                                            subName.put(instanceInfo.getNodeId(), instanceInfo.getNodeName())
                                            ObjectDefinitionNode upperNode = uppObjectModelAssigmentsName.get(it.getModelId())
                                            upperName.put(upperNode.getNodeId().getId(), getNodeFullName(upperNode))
                                            if (!uppSubId.containsKey(upperNode.getNodeId().id)) {
                                                uppSubId.put(upperNode.getNodeId().id, new ArrayList<String>())
                                            }
                                            uppSubId.get(upperNode.getNodeId().id).add(instanceInfo.nodeId)
                                        }
                                    }
                                }
                            }

                } else {
                    objectEntries = near.getObjectEntries().stream()
                            .filter { it.getModelTypeId().equalsIgnoreCase(EPC_ID) }
                            .collect(Collectors.toList())
                    if (objectEntries != null && !objectEntries.isEmpty() && objectEntries.stream()
                            .anyMatch { modelOfObject ->
                                modelOfObject.getObjectInstanceInfoList().stream()
                                        .anyMatch { it.getSymbolId().equalsIgnoreCase(ST_FUNC) || it.getSymbolId().equalsIgnoreCase(ST_SOLAR_FUNC) }
                            }) {
                        objectEntries.stream()
                                .filter { modelOfObject ->
                                    modelOfObject.getObjectInstanceInfoList().stream()
                                            .anyMatch { it.getSymbolId().equalsIgnoreCase(ST_FUNC) || it.getSymbolId().equalsIgnoreCase(ST_SOLAR_FUNC) }
                                }
                                .forEach {
                                    FullModelDefinition parent = getModelDefinition(modelApi, repositoryId, it.getModelId())
                                    parent.getModel().getParentNodesInfo().forEach { instanceInfo ->
                                        instanceInfo.getNodeAllocations().forEach { model ->
                                            if (uppObjectModelAssigmentsName.containsKey(model.getModelId())) {
                                                subName.put(instanceInfo.getNodeId(), instanceInfo.getNodeName())
                                                ObjectDefinitionNode upperNode = uppObjectModelAssigmentsName.get(model.getModelId())
                                                upperName.put(upperNode.getNodeId().getId(), getNodeFullName(upperNode))
                                                if (!uppSubId.containsKey(upperNode.getNodeId().id)) {
                                                    uppSubId.put(upperNode.getNodeId().id, new ArrayList<String>())
                                                }
                                                uppSubId.get(upperNode.getNodeId().id).add(instanceInfo.nodeId)
                                            } else if (it.getModelTypeId().equalsIgnoreCase(PSD_ID)) {
                                                FullModelDefinition parentParent = getModelDefinition(modelApi, repositoryId, model.getModelId())
                                                parentParent.getModel().getParentNodesInfo().forEach { instanceInfoParent ->
                                                    instanceInfoParent.getNodeAllocations().forEach { modelParent ->
                                                        if (uppObjectModelAssigmentsName.containsKey(modelParent.getModelId())) {
                                                            subName.put(instanceInfoParent.getNodeId(), instanceInfoParent.getNodeName())
                                                            ObjectDefinitionNode upperNode = uppObjectModelAssigmentsName.get(modelParent.getModelId())
                                                            upperName.put(upperNode.getNodeId().getId(), getNodeFullName(upperNode))
                                                            if (!uppSubId.containsKey(upperNode.getNodeId().id)) {
                                                                uppSubId.put(upperNode.getNodeId().id, new ArrayList<String>())
                                                            }
                                                            uppSubId.get(upperNode.getNodeId().id).add(instanceInfoParent.nodeId)
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                    }
                }
            }
        }

        if (!uppSubId.isEmpty()) {
            if (!isFirstLevel) {
                uppSubId.sort {upperName.get(it.getKey())}.each {uppId, subIds ->
                    func.add(BLT + upperName.get(uppId))
                    subIds.sort{subName.get(it)}.forEach {
                        func.add(BLT * 2 + subName.get(it))
                    }
                }
            } else {
                uppSubId.keySet().sort{upperName.get(it)}.stream().forEach {uppId ->
                    func.add(BLT + upperName.get(uppId))
                }
            }
        }

        return func.toList();
    }

    private static Map<Integer, List<String>> getGroupInfo(final Map<String, OrgUnitNode> rootNodes,
                                                           final Set<ObjectElementNode> sortedGroups,
                                                           final List<ObjectElementNode> orgUnitModel,
                                                           int appDocumentCount,
                                                           ModelApi modelApi,
                                                           String repositoryId) {
        List<String> groupInfo = new ArrayList<>()
        int appGroupCount = appDocumentCount + 1
        int groupCount = 1

        List<String> groupAndOrgUnit = new ArrayList<>()
        rootNodes.values().forEach {
            if (!it.isPosition) {
                setOrgUnitNameList(it, 1, groupAndOrgUnit, orgUnitModel)
            }
        }

        sortedGroups.forEach { groupObject ->
            {
                List<FullModelDefinition> groupDecomposition = groupObject.getObjectDefinitionNode().getModelAssignments()
                        .stream()
                        .map { getModelDefinition(modelApi, repositoryId, it.getModelId()) }
                        .collect(Collectors.toList())

                if (groupDecomposition != null && !groupDecomposition.isEmpty()) {
                    String groupName = groupObject.getName()
                    String groupHeader = new StringBuilder("\\B${appGroupCount}.${groupCount} СОСТАВ ГРУППЫ \"${groupName}\"")
                    String groupContent = new StringBuilder("\\AГруппа \"${groupName}\" включает:")

                    groupCount++
                    groupInfo.add(groupHeader)
                    groupInfo.add(groupContent)
                    groupInfo.add("\\AПодразделения:")
                    groupAndOrgUnit.add(BLT + groupName)

                    // Retrieve decomposition (child objects)
                    groupDecomposition.stream()
                            .filter { group -> group != null }
                            .forEach { group ->
                                group.getObjects().stream()
                                        .filter { it.getObjectTypeId() == 'OT_ORG_UNIT' }
                                        .map { g -> getNodeFullName(g) }
                                        .sorted(unitTreeComparator::compare)
                                        .forEach { fullName ->
                                            groupAndOrgUnit.add(BLT * 2 + fullName)
                                            groupInfo.add("${BLT}\\!B" + fullName)
                                        }
                            }
                }
            }
        }


        // If no groups were found, prevent further processing
        if (groupInfo.isEmpty()) {
            log.info("Объекты типа \"Группа\" не найдены. Переменная не будет заполнена.")
        }

        Map<Integer, List<String>> map = new HashMap<>();
        map.put(0, groupInfo);
        map.put(1, groupAndOrgUnit);
        return map;
    }

    private static void setOrgUnitNameList(OrgUnitNode node, int level, List<String> names, List<ObjectElementNode> orgUnits) {
        if (orgUnits.stream().anyMatch { node.id == it.id }) {
            names.add(BLT * level + node.getName())
        } else if (orgUnits.stream().anyMatch { node.containsInChild(it.getId()) }) {
            names.add(BLT * level + node.getName())
            try {
                node.getChildren().stream().sorted((n1, n2) -> unitTreeComparator.compare(n1.getName(), n2.getName()))
                        .forEach {
                            setOrgUnitNameList(it, level + 1, names, orgUnits)
                        }
            } catch (Exception e) {
                log.error(e.toString())
            }
        }
    }

    /**
     * Данные для раздела "Состав документа"
     */
    private List<String> getDocContent(Map<String, ObjectElementNode> functions, appDocumentCount, String presetId, String repositoryId) {
        ModelApi modelApi = context.getApi(ModelApi.class)
        EdgeTypeApi edgeTypeApi = context.getApi(EdgeTypeApi.class)

        Set<ObjectElementNode> sortedDocuments = new TreeSet<>((o1, o2) -> o1.getName() <=> o2.getName())

        for (ObjectElementNode fullDecomposition : functions.values()) {
            List<ObjectElementNode> elements = linkElements(fullDecomposition, INFO_CARR_ID, null, presetId, null, edgeTypeApi)

            // Check decomposition for "Information Carrier" objects
            elements.stream()
                    .filter {
                        it.getObjectInstance().getSymbolId().equalsIgnoreCase("ST_FOLD") ||
                                it.getObjectInstance().getSymbolId().equalsIgnoreCase("ST_CRD_FILE")
                    }
                    .forEach { sortedDocuments.add(it) }
        }


        int documentCount = 0
        ArrayList<String> documentInfo = new ArrayList<>()

        for (ObjectElementNode document : sortedDocuments) {
            List<String> subDocuments = new ArrayList<>()
            for (ModelAssignment modelAssignment : document.getObjectDefinitionNode().getModelAssignments()) {
                FullModelDefinition modelDefinition = getModelDefinition(modelApi, repositoryId, modelAssignment.modelId)
                if (modelDefinition != null && modelDefinition.getObjects() != null) {
                    modelDefinition.getObjects().stream()
                            .filter { it.nodeId.id != document.id }
                            .sorted(objectDefinitionNodeComparator)
                            .forEach {
                                subDocuments.add("${BLT}\\!B" + getNodeFullName(it))
                            }
                }
            }

            if (subDocuments.size() > 0) {
                String documentName = document.getName()
                documentCount++

                documentInfo.add("\\B${appDocumentCount}.${documentCount} СОСТАВ ДОКУМЕНТА \"${documentName}\"".toString())
                documentInfo.add("\\AДокумент \"$documentName\" включает следующие документы:".toString())
                documentInfo.addAll(subDocuments)
            }
        }

        // If no documents were found, prevent further processing
        if (documentCount == 0) {
            log.info("Объекты типа \"Носитель информации\" не найдены. Переменная не будет заполнена.")
        }

        return documentInfo
    }

    private void setChildrenNameList(ObjectElementNode parent, List<String> childrenList, List<ObjectElementNode> objectChildrenList, int level, Set<String> functionIds) {
        if (parent == null) {
            return
        }

        // Формирование строки с отступами для родителя
        String indent = BLT * level

        // Добавление имени родителя в список, если условия выполнены
        if (level == 1 && shouldAddElement(parent, functionIds)) {
            childrenList.add(indent + parent.getName())
            objectChildrenList.add(parent)
        }

        // Проверка и обработка потомков
        if (parent.children != null && !parent.children.isEmpty() && hasPathToFunction(parent, functionIds)) {
            int newLevel = level + 1
            Comparator<ObjectElementNode> comparator = CHILD_COMPARATOR
            if (parent.children.stream().anyMatch { child -> child.getModelDefinition().getModel().getModelTypeId().equalsIgnoreCase(PSD_ID) }) {
                comparator = PSD_COMPARATOR
            }
            parent.children.stream()
                    .sorted(comparator)
                    .forEach(child -> processChild(child, childrenList, objectChildrenList, newLevel, functionIds))
        }
    }

    private boolean shouldAddElement(ObjectElementNode element, Set<String> functionIds) {
        String symbolId = element.getObjectInstance().getSymbolId()
        return hasPathToFunction(element, functionIds) && !stWithout.contains(symbolId.toUpperCase())
    }

    private void processChild(ObjectElementNode child, List<String> childrenList, List<ObjectElementNode> objectChildrenList, int level, Set<String> functionIds) {
        // Добавление ребенка в список, если условия выполнены
        if (shouldAddElement(child, functionIds)) {
            String indent = BLT * level
            childrenList.add(indent + child.getName())
            objectChildrenList.add(child)
        }

        // Рекурсивная обработка потомков
        if (child.children != null && !child.children.isEmpty()) {
            setChildrenNameList(child, childrenList, objectChildrenList, level, functionIds)
        }
    }

    private boolean hasPathToFunction(ObjectElementNode parent, Set<String> functionIds) {
        if (parent == null) {
            return false
        }

        boolean hasMatchingDescendant = false

        // Проверяем наличие потомков
        if (parent.children != null && parent.children.size() > 0) {
            for (ObjectElementNode child : parent.children) {
                // Рекурсивно проверяем потомков
                boolean childHasMatch = hasPathToFunction(child, functionIds)

                // Если хотя бы один потомок ведет к функции, отмечаем это
                if (childHasMatch) {
                    hasMatchingDescendant = true
                }
            }
        }

        // Возвращаем true, если родитель является функцией или хотя бы один потомок ведет к функции
        return (functionIds.contains(parent.id) && parent.getParent() != null) || hasMatchingDescendant
    }

    @SuppressWarnings("unchecked")
    private static void loadCacheFromFile() {
        try (ObjectInputStream ois = new ObjectInputStream(new FileInputStream("${LOCAL_PATH}\\cacheData.ser"))) {
            // Загружаем каждый кэш
            cacheFullModelDefinition = (Map<String, FullModelDefinition>) ois.readObject()
            cacheObjectDefinition = (Map<String, ObjectDefinitionNode>) ois.readObject()
            sourceCache = (Map<String, ObjectElementNode>) ois.readObject()
            targetCache = (Map<String, ObjectElementNode>) ois.readObject()
            bulletNumberingMap = (Map<Integer, BigInteger>) ois.readObject()
            linkTypeNameCache = (Map<String, String>) ois.readObject()
        } catch (Exception e) {
            // Файл не существует, можно проигнорировать, так как это первый запуск
            log.error(e.toString())
        }
    }

    private static void saveCacheToFile() {
        Path filePath = Paths.get("${LOCAL_PATH}\\cacheData.ser")

        try {
            // Удаляем файл, если он существует
            Files.deleteIfExists(filePath)

            // Создаём новый файл и записываем данные
            try (ObjectOutputStream oos = new ObjectOutputStream(new FileOutputStream(filePath.toFile()))) {
                // Сохраняем каждый кэш
                oos.writeObject(cacheFullModelDefinition)
                oos.writeObject(cacheObjectDefinition)
                //oos.writeObject(linkElementsEntriesCache)
                oos.writeObject(sourceCache)
                oos.writeObject(targetCache)
                oos.writeObject(bulletNumberingMap)
                oos.writeObject(linkTypeNameCache)
            }
        } catch (Exception e) {
            // Файл не существует, можно проигнорировать, так как это первый запуск
            log.error(e.toString())
        }
    }

    private String getStatus(ObjectElementNode element, EdgeTypeApi edgeTypeApi, String presetId) {
        ObjectElementNode status = linkElements(element, OT_EVT, null, presetId, null, edgeTypeApi).stream()
                .filter { it.getObjectInstance().getSymbolId() == STATUS_SYMBOL_ID }.findFirst().orElse(null)


        // Если атрибут не найден или не заполнен, возвращаем пустую строку
        return status != null ? status.getName() : ""
    }

    private static Map<String, List<List<String>>> fillNewFunctionTable(Map<String, Map<String, Object>> functionDataMap, Set<ObjectElementNode> subProcessRoles, List<String> sortSubProcess) {
        Map<String, List<List<String>>> createTablesToData = new LinkedHashMap<>()
        List<String> header = ["\\B№ п/п", "\\BНаименование операции", "\\BДокументы", "\\BОтветственный (бизнес-роль, подразделение)", "\\BРезультат операции", "\\BИнформационная система", "\\BКомментарий"]

        // Обработка для каждого субпроцесса
        List<ObjectElementNode> sortSubProcessRoles = subProcessRoles.stream().collect(Collectors.toList())
        Comparator<ObjectElementNode> comparator = Comparator.comparing({ ObjectElementNode node ->
            {
                String subProc = sortSubProcess.find { sub -> sub.containsIgnoreCase(node.getName()) }
                return subProc == null ? -1 : sortSubProcess.indexOf(subProc)
            }
        })
        Collections.sort(sortSubProcessRoles, comparator)
        sortSubProcessRoles.stream().forEach(subProcessRole -> {
            // Инициализация таблицы для каждого процесса отдельно
            List<List<String>> tableData = new ArrayList<>()
            boolean createTable = false
            // Обрабатываем каждую функцию
            List<ObjectElementNode> children = subProcessRole.children.stream().collect(Collectors.toList())
            Collections.sort(children, CHILD_COMPARATOR)
            int counter = 1
            for (ObjectElementNode function : children) {

                Map<String, Object> functionData = functionDataMap.get(function.id)

                if (functionData != null && functionData.size() > 0) {
                    createTable = true

                    // Создаем строку для таблицы
                    List<String> row = new ArrayList<>()
                    row.add(String.valueOf(counter++)) // № п/п
                    row.add((String) functionData.get("functionName")) // Наименование операции (Функция)

                    // Обработка входящих документов
                    Map<String, String> incomingDocsWithStatus = (Map<String, String>) functionData.get("incomingDocsWithStatus")
                    String incomingDocsText = incomingDocsWithStatus != null ?
                            incomingDocsWithStatus.entrySet().stream()
                                    .map(entry ->
                                            entry.getValue() != null && !entry.getValue().isEmpty()
                                                    && trimName(entry.getValue()) != "" ?
                                                    entry.getKey() + " (статус: " + entry.getValue() + ")" : entry.getKey())
                                    .collect(Collectors.joining(",/n")) : ""

                    // Обработка исходящих документов
                    Map<String, String> outgoingDocsWithStatus = (Map<String, String>) functionData.get("outgoingDocsWithStatus")
                    String outgoingDocsText = outgoingDocsWithStatus != null ?
                            outgoingDocsWithStatus.entrySet().stream()
                                    .map(entry ->
                                            entry.getValue() != null && !entry.getValue().isEmpty()
                                                    && trimName(entry.getValue()) != "" ?
                                                    entry.getKey() + " (статус: " + entry.getValue() + ")" : entry.getKey())
                                    .collect(Collectors.joining(",/n")) : ""

                    incomingDocsText = incomingDocsText != null && !incomingDocsText.isEmpty() ? "\\BВходящие документы:/n" + incomingDocsText : ""
                    outgoingDocsText = outgoingDocsText != null && !outgoingDocsText.isEmpty() ? "/n\\BИсходящие документы:/n" + outgoingDocsText : ""

                    row.add(incomingDocsText + " " + outgoingDocsText)

                    // Ответственный (бизнес-роль)
                    List<String> responsibleList = (List<String>) functionData.get("responsible")
                    String responsibleText = responsibleList != null ?
                            responsibleList.stream()
                                    .map(responsible -> "\\BВыполняет:/n " + responsible)
                                    .collect(Collectors.joining(", /n")) : ""
                    row.add(responsibleText)

                    // Результат операции
                    List<String> results = (List<String>) functionData.get("results")
                    String resultText = results != null ?
                            results.stream().collect(Collectors.joining("/n")) : ""
                    row.add(resultText)

                    // Информационная система
                    List<String> systems = (List<String>) functionData.get("systems")
                    String systemText = systems != null ?
                            systems.stream().collect(Collectors.joining(", ")) : ""
                    row.add(systemText)

                    // Комментарий (пустое поле)
                    row.add("")

                    tableData.add(row)
                }
            }

            if (createTable) {
                tableData.add(0, header)
                // Добавляем данные в map для конкретного субпроцесса
                createTablesToData.put(subProcessRole.getName(), tableData)
            }
        })

        return createTablesToData
    }


    List<String> getOrganizationalUnitsAndInParameterLists(Map<String, OrgUnitNode> rootNodes, Map<String, OrgUnitNode> visited, Set<ObjectElementNode> sortedGroups, int appCount, String presetId) {
        ObjectsApi objectsApi = context.getApi(ObjectsApi.class)
        // Список строк для заполнения тега <Подразделение>
        List<String> subdivisionsList = new ArrayList<>()

        sortedGroups.forEach {
            List<ObjectElementNode> groupUnit = new ArrayList<>()
            groupUnit.addAll(linkElementsEntries(it, OT_ORG_UNIT, null, null, presetId, null))
            if (groupUnit != null && !groupUnit.isEmpty()) {
                for (ObjectElementNode businessRole : groupUnit) {
                    String orgUnitId = businessRole.getObjectDefinitionNode().nodeId.id

                    OrgUnitNode orgUnitNode
                    if (visited.containsKey(orgUnitId)) {
                        // Если узел уже был посещен, используем его
                        orgUnitNode = visited.get(orgUnitId)
                    } else {
                        // Иначе создаем новый узел
                        orgUnitNode = new OrgUnitNode(businessRole.getObjectDefinitionNode(), false)
                        visited.put(orgUnitNode.id, orgUnitNode)
                    }
                    addPositionToTree(rootNodes, orgUnitNode, objectsApi)
                }
            }
        }

        // Пройдемся по всем подразделениям и сформируем строки
        rootNodes.values().each {
            traverseOrgUnitTree(it, 1, subdivisionsList)
        }

        int i = 1
        sortedGroups.forEach {
            String text = it.getName()
            subdivisionsList.add("${BLT}${text}(состав Группы см. п.${i} Приложение ${appCount})".toString())
            i++
        }

        // Записываем в parameterLists
        return subdivisionsList
    }

    // Заполнение таблицы с использованием всех бизнес-ролей и организационных единиц
    List<List<String>> fillOrganizationalUnitsTable(Map<String, OrgUnitNode> rootNodes) {
        // Данные для таблицы
        List<List<String>> tableData = []

        // Общий корневой узел для всех функций
        rootNodes.values().stream().filter { !it.isPosition }.forEach { rootNode ->
            buildTableDataFromTree(rootNode, 0, tableData)
        }

        // #76596 Для п.3 Таблицы 1, добавить Бизнес-роли без Должности
        Set<String> orphanBusinessRoles = rootNodes.values().stream()
                .filter { rn -> rn.isPosition }
                .flatMap { rn -> rn.getBusinessRoles().stream() }
                .collect(Collectors.toSet())

        for (String role : orphanBusinessRoles.sort()) {
            List<String> row = [role]
            tableData.add(row)
        }

        // Добавляем данные в numberTablesToData
        return tableData // Индекс таблицы может быть изменен в зависимости от структуры документа
    }

    /**
     * Формирует таблицу "Функциональные роли участников процесса"
     * @param businessRolesForFunctions
     * @param parentProcessForFunction
     * @param functions
     * @param isFirstLevel
     * @return
     */
    List<List<String>> fillFunctionalRolesTable(Map<String, List<String>> businessRolesForFunctions, Map<String, String> parentProcessForFunction, Map<String, ObjectElementNode> functions, boolean isFirstLevel) {
        // Данные для таблицы
        List<List<String>> tableData = []

        // Заполняем строки для каждой функции
        functions.each { functionId, function ->
            Set<String> businessRoles = businessRolesForFunctions.get(functionId).stream().collect(Collectors.toSet())
            String parentProcessName = parentProcessForFunction.get(functionId) ?: ""

            for (String roleName : businessRoles) {
                // Заполняем строку таблицы
                List<String> row = new ArrayList<>()
                row.add(roleName) // Бизнес-роль в первой колонке
                if (isFirstLevel) {
                    row.add(function.getRoot().getName())
                    row.add(parentProcessName)
                } else {
                    row.add(parentProcessName) // Родительский процесс
                    row.add(function.getName()) // Имя функции как наименование процедуры
                }

                // Добавляем строку в данные для таблицы
                tableData.add(row)
            }
        }

        // Сортировка по трем столбцам: #77011
        tableData.sort { "${it.get(0).trim()} ${it.get(1).trim()} ${it.get(2).trim()}".toLowerCase() }
        return tableData // Индекс может быть изменен в зависимости от вашей структуры документа
    }

    // Метод для объединения деревьев без дублирования узлов
    void mergeOrgUnitTree(OrgUnitNode existingNode, OrgUnitNode newNode, Map<String, OrgUnitNode> nodeMap) {
        // Если новый узел уже существует в основном дереве
        if (nodeMap.containsKey(newNode.id)) {
            OrgUnitNode existingExistingNode = nodeMap.get(newNode.id)

            // Объединяем бизнес-роли, если узел является должностью
            if (existingExistingNode.isPosition && newNode.isPosition) {
                existingExistingNode.businessRoles.addAll(newNode.businessRoles)
            }

            // Рекурсивно объединяем детей
            newNode.children.each { child ->
                mergeOrgUnitTree(existingExistingNode, child, nodeMap)
            }
        } else {
            // Если узел не существует, добавляем его в дерево и карту
            existingNode.addChild(newNode)
            nodeMap.put(newNode.id, newNode)

            // Рекурсивно добавляем всех детей
            newNode.children.each { child ->
                mergeOrgUnitTree(newNode, child, nodeMap)
            }
        }
    }

    // Метод для построения данных таблицы из дерева
    void buildTableDataFromTree(OrgUnitNode node, int level, List<List<String>> tableData) {
        if (node.name != "") {
            String indent = "\\L${level}"
            String orgUnitText = indent + node.name

            if (node.isPosition) {
                // Узел представляет должность
                if (!node.businessRoles.isEmpty()) {
                    // Формируем маркированный список бизнес-ролей
                    String businessRolesText = node.businessRoles.sort().collect { role -> BLT + role }.join("/n")
                    List<String> row = [orgUnitText, businessRolesText]
                    tableData.add(row)
                }
            } else {
                // Узел представляет организационную единицу
                List<String> row = [orgUnitText, ""]
                tableData.add(row)
            }
        }

        // Рекурсивно обрабатываем детей
        node.children.sort { it.name.toLowerCase() }.each { child ->
            buildTableDataFromTree(child, level + 1, tableData)
        }
    }


    List<ObjectElementNode> linkElementsEntries(ObjectElementNode el, String typeId, String modelType, String linkType, String presetId, String connectionType) {
        ModelApi modelApi = context.getApi(ModelApi.class)
        EdgeTypeApi edgeTypeApi = context.getApi(EdgeTypeApi.class)
        ObjectsApi objectsApi = context.getApi(ObjectsApi.class)

        // Формируем уникальный ключ для кэша на основе параметров
        String cacheKey = el.getObjectDefinitionNode().getNodeId().getId() + "_" + typeId + "_" + modelType + "_" + linkType + "_" + connectionType

        // Проверяем кэш
        if (linkElementsEntriesCache.containsKey(cacheKey)) {
            return linkElementsEntriesCache.get(cacheKey)
        }

        List<ObjectElementNode> result = new ArrayList<>()
        String objectNodeId = el.getObjectDefinitionNode().getNodeId().getId()

        // Получаем ObjectDefinitionNode из кэша или загружаем
        ObjectDefinitionNode objectNode
        try {
            objectNode = getObjectDefinition(objectsApi, el.getObjectDefinitionNode().getNodeId().getRepositoryId(), objectNodeId)
        } catch (Exception e) {
            log.error("Ошибка при получении ObjectDefinitionNode: ", e)
            return Collections.emptyList() // Возвращаем пустой результат в случае ошибки
        }

        objectNode.getObjectEntries().stream().forEach { parent ->
            if (modelType == null || parent.getModelTypeId().equalsIgnoreCase(modelType)) {
                String modelId = parent.getModelId()

                // Получаем модель из кэша
                FullModelDefinition modelDefinition
                try {
                    modelDefinition = getModelDefinition(modelApi, objectNode.getNodeId().repositoryId, modelId)
                } catch (Exception e) {
                    log.error("Ошибка при получении FullModelDefinition: ", e)
                    return // Продолжаем обработку других элементов, если произошла ошибка
                }
                if (modelDefinition != null) {
                    getObjectElementsNode(modelDefinition, objectNode, null, edgeTypeApi, presetId).forEach {
                        result.addAll(linkElements(it, typeId, linkType, presetId, connectionType, edgeTypeApi))
                    }
                }
            }
        }

        // Сохраняем результат в кэш
        linkElementsEntriesCache.put(cacheKey, result)
        return result
    }

    private static List<ObjectElementNode> linkElements(ObjectElementNode el, String typeId, String linkType, String presetId, String connectionType, EdgeTypeApi edgeTypeApi) {
        List<ObjectElementNode> result = new ArrayList<>()
        if (el == null) {
            return result
        }

        for (int i = 0; i < 2; i++) {
            // Проверяем тип связи на входящая/исходящая или обе (null)
            if ((connectionType == null) ||
                    (connectionType.equals("incoming") && i == 0) ||
                    (connectionType.equals("outgoing") && i == 1)) {

                List<EdgeInstance> edges = i == 0 ? el.getEnterEdges() : el.getExitEdges()
                for (EdgeInstance edge : edges) {
                    // TODO распутать и оптимизировать логику
                    ObjectElementNode linkedElement = i == 0 ? el.getSource(edge) : el.getTarget(edge)
                    if (linkedElement != null && linkedElement.getObjectDefinitionNode().getObjectTypeId().equalsIgnoreCase(typeId)) {
                        if (linkType == null || linkTypeNameCache.computeIfAbsent(edge.edgeTypeId, edgeTypeId ->
                                edgeTypeApi.byId(presetId, edgeTypeId).getName().toLowerCase()
                        ).equals(linkType.toLowerCase())) {
                            result.add(linkedElement)
                        }
                    }
                }
            }
        }
        return result
    }

    private static List<ObjectElement> linkElements(ObjectElement el, String typeId, String linkType, String presetId, EdgeTypeApi edgeTypeApi) {
        List<ObjectElement> result = new ArrayList<>()
        if (el == null) {
            return result
        }
        for (int i = 0; i < 2; i++) {
            List<Edge> edges = i == 0 ? el.getEnterEdges() : el.getExitEdges()
            for (Edge edge : edges) {
                ObjectElement linkedElement = (i == 0 ? edge.getSource() : edge.getTarget()) as ObjectElement
                if (linkedElement != null) {
                    ObjectDefinition objectDefinition = linkedElement.getObjectDefinition()
                    if (objectDefinition != null && objectDefinition.getObjectTypeId().toLowerCase().equals(typeId.toLowerCase())) {
                        if (linkType == null ||
                                edgeTypeApi.byId(presetId, edge.edgeTypeId).getName().toLowerCase().equals(linkType.toLowerCase())) {
                            result.add(linkedElement)
                        }
                    }
                }
            }
        }
        return result
    }

    private static List<String> linkElementsName(ObjectElement el, String typeId, String linkType, String presetId, EdgeTypeApi edgeTypeApi) {
        return linkElements(el, typeId, linkType, presetId, edgeTypeApi)
                .stream()
                .map { it -> it.getObjectDefinition().getName() }
                .collect(Collectors.toList())
    }

    private static List<ObjectElementNode> getObjectElementsNode(FullModelDefinition fullModel, ObjectDefinitionNode objectNode, ObjectElementNode parent, EdgeTypeApi edgeTypeApi, String presetId) {
        List<ObjectElementNode> elements = new ArrayList<>()
        List<String> elementIds = fullModel.model.getElements().stream()
                .filter { it.type == DiagramElementType.object }
                .map { ObjectInstance.cast(it) }.filter { it.getObjectDefinitionId() == objectNode.getNodeId().id }.map { it.getId() }.collect(Collectors.toList())
        elementIds.forEach {
            elements.add(new ObjectElementNode(fullModel, objectNode, parent, it, edgeTypeApi, presetId))
        }

        return elements
    }


    /**
     * Выборка шаблона
     * @param templateName - имя шаблона
     * @return шаблон
     */
    Map<String, ObjectElementNode> findFunctionByLevelDecomposition(ObjectElementNode parent,
                                                                    Map<String, ObjectElementNode> functionsAll,
                                                                    Map<String, ObjectDefinitionNode> nearProcessNode,
                                                                    FullModelDefinition contextModel,
                                                                    String presetId,
                                                                    int maxDepth,
                                                                    String repositoryId,
                                                                    TreeRepository repository,
                                                                    boolean isFirstLevel,
                                                                    String parentId) {
        ModelApi modelApi = context.getApi(ModelApi.class)
        EdgeTypeApi edgeTypeApi = context.getApi(EdgeTypeApi.class)
        ObjectTypeApi objectTypeApi = context.getApi(ObjectTypeApi.class)
        ObjectsApi objectsApi = context.getApi(ObjectsApi.class)

        Map<String, ObjectElementNode> functions = new HashMap<>()
        findFunctionByLevelDecomposition(functionsAll, contextModel, parent, presetId,
                maxDepth, edgeTypeApi, objectTypeApi, 2, functions, nearProcessNode, modelApi, objectsApi, repositoryId, repository, isFirstLevel, parentId)
        return functions
    }

    static void findFunctionByLevelDecomposition(Map<String, ObjectElementNode> functionsAll, FullModelDefinition model, ObjectElementNode parentNode, String presetId,
                                                 int maxDepth, EdgeTypeApi edgeTypeApi,
                                                 ObjectTypeApi objectTypeApi, int currentDepth, Map<String, ObjectElementNode> functions, Map<String, ObjectDefinitionNode> nearProcessNode,
                                                 ModelApi modelApi, ObjectsApi objectsApi, String repositoryId, TreeRepository repository, boolean isFirstLevel, String parentId) {

        // Получаем список объектов модели
        List<ObjectDefinitionNode> objects = model.getObjects()
        int step = 1

        // Список для отложенной обработки декомпозиций
        List<ObjectElementNode> deferredDecompositions = new ArrayList<>()
        List<ObjectDefinitionNode> processFun = new ArrayList<>()

        if (model.getModel().getModelTypeId().equalsIgnoreCase(PSD_ID)) {
            step = 0
        }

        if (currentDepth > detailLevel && step != 0) {
            // Если текущая глубина больше максимальной, выходим из рекурсии
            return
        }

        // Проходим по объектам на текущем уровне
        objects.forEach { obj ->
            try {
                if (!isFirstLevel && parentId != null && !obj.getNodeId().id.equalsIgnoreCase(parentId)) {
                    return
                }

                // Проверяем, является ли объект функцией
                ObjectInstance instance = model.model.getElements().stream()
                        .filter { it.type == DiagramElementType.object }
                        .map { ObjectInstance.cast(it) }.filter { it -> it.getObjectDefinitionId() == obj.getNodeId().getId() }.findFirst().orElse(null)
                String symbolId = instance.getSymbolId()

                if (FUNCTION_ID.equalsIgnoreCase(obj.getObjectTypeId())) {
                    // Добавляем в список смежных процессов для дальнейшей обработки
                    if (!(currentDepth < maxDepth) && symbolId.equalsIgnoreCase(REL_I_PROCESS)) {
                        nearProcessNode.put(obj.getNodeId().id, getObjectDefinition(objectsApi, repositoryId, obj.getNodeId().id))

                        return
                    }

                    if (stWithout.contains(instance.getSymbolId().toUpperCase())) {
                        return
                    }

                    if (functionsAll.containsKey(obj.getNodeId().id)) {
                        return
                    }

                    if (symbolId.equalsIgnoreCase(ST_PRCS_1)) {
                        if (step == 0) {
                            processFun.add(obj)
                        }

                        return
                    }

                    getObjectElementsNode(model, obj, parentNode, edgeTypeApi, presetId).forEach { elementNode ->
                        {
                            // Создаем элемент узла для текущего объекта
                            functionsAll.put(elementNode.getId(), elementNode)

                            // Если функция должна быть добавлена на этом уровне
                            if (!(currentDepth < maxDepth) && step != 0) {
                                functions.put(elementNode.id, elementNode)
                            }

                            // Сохраняем декомпозиции для последующей обработки
                            Set<ModelAssignment> assignments = obj.getModelAssignments()
                            if (assignments != null && !assignments.isEmpty() && currentDepth <= maxDepth) {
                                deferredDecompositions.add(elementNode) // добавляем узел в список для декомпозиций
                            }
                        }
                    }

                }
            } catch (Exception e) {
                log.error("Ошибка декомпозиции для объекта ${getNodeFullName(obj)}: ", e)
            }
        }

        if (step == 0 && deferredDecompositions.isEmpty()) {
            processFun.each {
                getObjectElementsNode(model, it, parentNode, edgeTypeApi, presetId).forEach { process ->
                    functions.put(process.getId(), process)
                    functionsAll.put(process.getId(), process)
                }
            }
            return
        }

        // Теперь обрабатываем отложенные декомпозиции, только после обработки всех объектов текущего уровня
        deferredDecompositions.each { elementNode ->
            {
                Set<ModelAssignment> assignments = elementNode.getObjectDefinitionNode().getModelAssignments()
                assignments.each { assignment ->
                    {
                        if (!assignment.getModelTypeId().equalsIgnoreCase(PSD_ID) && !assignment.getModelTypeId().equalsIgnoreCase(EPC_ID)) {
                            return
                        }
                        if (!assignment.modelId.equals(model.getModel().getNodeId().id)) {
                            long start = System.currentTimeMillis()
                            FullModelDefinition fullModelDefinition = getModelDefinition(modelApi, repositoryId, assignment.getModelId())
                            log.debug("Время выполнения getModelDefinition = ${System.currentTimeMillis() - start}")

                            // Рекурсивно переходим на следующий уровень декомпозиции
                            if (fullModelDefinition != null) {
                                findFunctionByLevelDecomposition(
                                        functionsAll,
                                        fullModelDefinition,
                                        elementNode,
                                        presetId, maxDepth, edgeTypeApi, objectTypeApi,
                                        currentDepth + step, functions, nearProcessNode, modelApi, objectsApi, repositoryId, repository, false, null
                                )
                            }
                        }
                    }
                }
            }
        }
    }

    static Set<ObjectElementNode> findGroup(FullModelDefinition model, EdgeTypeApi edgeTypeApi, String presetId) {
        Set<ObjectElementNode> groups = new HashSet<>()

        if (model != null) {
            List<ObjectDefinitionNode> objects = model.getObjects()
            objects.forEach { object ->
                try {
                    if (GROUP_ID.equalsIgnoreCase(object.getObjectTypeId())) {
                        getObjectElementsNode(model, object, null, edgeTypeApi, presetId).forEach {
                            groups.add(it)
                        }
                    }
                } catch (Exception e) {
                    log.error(e.toString())
                }
            }
        }

        return groups

    }

    static void findDecomposition(FullModelDefinition model, ObjectElementNode parentNode, String presetId, EdgeTypeApi edgeTypeApi,
                                  ObjectTypeApi objectTypeApi, Map<String, FullModelDefinition> models,
                                  ModelApi modelApi, String repositoryId, boolean isFirstLevel, String parentId, int detailLevel) {

        // Получаем список объектов модели
        List<ObjectDefinitionNode> objects = model.getObjects()

        // Список для отложенной обработки декомпозиций
        List<ObjectElementNode> deferredDecompositions = new ArrayList<>()

        /*if (model.getModel().getModelTypeId().equalsIgnoreCase(PSD_ID)) {
            step = 0
        }
*/
        // Проходим по объектам на текущем уровне
        objects.each { obj ->
            try {
                if (!isFirstLevel && parentId != null && !obj.getNodeId().id.equalsIgnoreCase(parentId)) {
                    return
                }

                // Проверяем, является ли объект функцией
                ObjectInstance instance = model.model.getElements().stream()
                        .filter { it.type == DiagramElementType.object }
                        .map { ObjectInstance.cast(it) }.filter { it -> it.getObjectDefinitionId() == obj.getNodeId().getId() }.findFirst().orElse(null)
                String symbolId = instance.getSymbolId()

                if (stWithout.contains(instance.getSymbolId().toUpperCase())) {
                    return
                }

                // Проверяем, является ли объект функцией
                if (FUNCTION_ID.equalsIgnoreCase(obj.getObjectTypeId())) {
                    if (model.getModel().getModelTypeId().equalsIgnoreCase(PSD_ID) && symbolId.equalsIgnoreCase("ST_PRCS_1")) {
                        return
                    }

                    // Сохраняем декомпозиции для последующей обработки
                    Set<ModelAssignment> assignments = obj.getModelAssignments()
                    if (assignments != null && !assignments.isEmpty()) {
                        getObjectElementsNode(model, obj, parentNode, edgeTypeApi, presetId).forEach { elementNode ->
                            deferredDecompositions.add(elementNode) // добавляем узел в список для декомпозиций
                        }
                    }
                }
            } catch (Exception e) {
                log.error("Ошибка декомпозиции для объекта ${getNodeFullName(obj)}: ", e)
            }
        }

        /*if (model.getModel().getModelTypeId().equalsIgnoreCase(PSD_ID) && deferredDecompositions.isEmpty()) {
            models.put(model.getModel().getNodeId().id, model)
        }*/

        // Теперь обрабатываем отложенные декомпозиции, только после обработки всех объектов текущего уровня
        deferredDecompositions.stream().sorted().forEach { elementNode ->
            {
                Set<ModelAssignment> assignments = elementNode.getObjectDefinitionNode().getModelAssignments()
                assignments.each { assignment ->
                    {
                        if (!assignment.getModelTypeId().equalsIgnoreCase(PSD_ID) && !assignment.getModelTypeId().equalsIgnoreCase(EPC_ID)) {
                            return
                        }
                        if (!assignment.modelId.equals(model.getModel().getNodeId().id)) {
                            int step = 1
                            long start = System.currentTimeMillis()
                            FullModelDefinition fullModelDefinition = getModelDefinition(modelApi, repositoryId, assignment.getModelId())
                            if (fullModelDefinition == null) {
                                return
                            } else if (fullModelDefinition.getModel().getModelTypeId().equalsIgnoreCase(PSD_ID)) {
                                step = 0
                            }
                            models.put(fullModelDefinition.getModel().getNodeId().id, fullModelDefinition)
                            log.debug("Время выполнения getModelDefinition = ${System.currentTimeMillis() - start}")

                            // Рекурсивно переходим на следующий уровень декомпозиции
                            if (detailLevel == 4 || step == 0) {
                                detailLevel = step == 0 ? detailLevel : 0
                                findDecomposition(
                                        fullModelDefinition,
                                        elementNode,
                                        presetId, edgeTypeApi, objectTypeApi, models, modelApi, repositoryId, false, null, detailLevel
                                )
                            }
                        }
                    }
                }
            }
        }
    }

    static class DepartmentNode {
        String name
        List<DepartmentNode> children = new ArrayList<>()

        DepartmentNode(String name) {
            this.name = name
        }
    }

    static class ObjectElementNode implements Comparable, Serializable {
        String id
        transient EdgeTypeApi edgeTypeApi
        String presetId
        FullModelDefinition modelDefinition // модель, на которой находится objectInstance
        ObjectDefinitionNode objectDefinitionNode
        List<EdgeInstance> edgeInstances
        ObjectInstance objectInstance
        List<ObjectInstance> objectInstances
        ObjectElementNode parent
        HashSet<ObjectElementNode> children = new HashSet<>()
        ObjectElementNode group
        Double x
        Double y

        ObjectElementNode(FullModelDefinition modelDefinition, ObjectDefinitionNode objectDefinitionNode, ObjectElementNode parent, String elementId, EdgeTypeApi edgeTypeApi, String presetId) {
            this.id = objectDefinitionNode.getNodeId().getId()
            this.modelDefinition = modelDefinition
            this.objectDefinitionNode = objectDefinitionNode
            this.edgeInstances = modelDefinition.model.getElements().stream()
                    .filter { it.type == DiagramElementType.edge }
                    .map { EdgeInstance.cast(it) }.collect(Collectors.toList())
            this.objectInstances = modelDefinition.model.getElements().stream()
                    .filter { it.type == DiagramElementType.object }
                    .map { ObjectInstance.cast(it) }.collect(Collectors.toList())
            this.objectInstance = this.objectInstances.find { it.getId() == elementId }
            this.parent = parent
            if (parent != null) {
                parent.addChild(this)
            }
            this.x = objectInstance.getX()
            this.y = objectInstance.getY()
            this.edgeTypeApi = edgeTypeApi
            this.presetId = presetId
        }

        List<EdgeInstance> getEnterEdges() {
            return getEdges(true)
        }

        List<EdgeInstance> getExitEdges() {
            return getEdges(false)
        }

        private List<EdgeInstance> getEdges(boolean isEnter) {
            return this.edgeInstances.stream()
                    .filter { isEnter ? it.target == objectInstance.getId() : it.source == objectInstance.getId() }
                    .collect(Collectors.toList())
        }

        @Description("Получить объект источник связи")
        ObjectElementNode getSource(EdgeInstance edgeDefinition) {
            return sourceCache.computeIfAbsent(edgeDefinition.source, { key ->
                new ObjectElementNode(modelDefinition, modelDefinition.getObjects().find { it.nodeId.id == objectInstances.find { it.id == key }.objectDefinitionId }, null, key, edgeTypeApi, presetId)
            })
        }

        @Description("Получить объект цель связи")
        ObjectElementNode getTarget(EdgeInstance edgeDefinition) {
            return targetCache.computeIfAbsent(edgeDefinition.target, { key ->
                new ObjectElementNode(modelDefinition, modelDefinition.getObjects().find { it.nodeId.id == objectInstances.find { it.id == key }.objectDefinitionId }, null, key, edgeTypeApi, presetId)
            })
        }

        ObjectElementNode getRoot() {
            if (parent == null) {
                return this
            }

            return parent.getRoot()
        }

        List<ObjectElementNode> getPath() {
            List<ObjectElementNode> path1 = new ArrayList<>()
            path1.add(this)

            ObjectElementNode parent = this.getParent()
            while (parent != null) {
                path1.add(parent)
                parent = parent.getParent()
            }

            return path1.reverse(true)
        }

        ObjectElementNode addChild(ObjectElementNode child) {
            children.add(child)
            return this
        }

        ObjectElementNode getGroup() {
            if (group == null) {
                group = linkElements(this, FUNCTION_ID, null, presetId, null, edgeTypeApi).stream().filter { it.getObjectInstance().getSymbolId().equalsIgnoreCase(ST_GROUP_1_ID) }.findFirst().orElse(null)
            }
            return group
        }

        String getName() {
            return getNodeFullName(objectDefinitionNode)
        }

        @Override
        public int compareTo(Object o) {
            if (o == null || !(o instanceof ObjectElementNode)) {
                throw new ClassCastException("Invalid object for comparison")
            }

            ObjectElementNode other = (ObjectElementNode) o

            return Comparator.comparing({ s -> (getGroup() != null) ? getGroup().getY() : 0.0d })
                    .thenComparing(Comparator.comparing({ x -> ((ObjectElementNode) x).getX() }))
                    .thenComparing(Comparator.comparing({ y -> ((ObjectElementNode) y).getY() })).compare(this, other)
        }

        boolean equals(o) {
            if (this.is(o)) return true
            if (o == null || getClass() != o.class) return false

            ObjectElementNode that = (ObjectElementNode) o

            if (id != that.id) return false

            return true
        }

        int hashCode() {
            return (id != null ? id.hashCode() : 0)
        }
    }

    static class ProcessPng implements Serializable {
        static final long serialVersionUID = 42L

        String name
        byte[] png

        ProcessPng() {
        }

        ProcessPng(String name, byte[] png) {
            this.name = name
            this.png = png
        }
    }

    private TreeNode findFolder(TreeRepository repository, String repositoryId, String folderId, String name) {
        SearchApi searchApi = context.getApi(SearchApi.class)

        SearchRequest request = new SearchRequest()
        request.setRootSearchNodeId(NodeId.builder().id(repositoryId).repositoryId(folderId).build())
        request.setSearchText(name)
        request.setSearchVisibility(SearchVisibility.NOT_DELETED)

        def searchResults = searchApi.searchExtended(request).getResultList()

        SearchResult sr = searchResults.stream()
                .filter { it.nodeType == NodeType.FOLDER }
                .findFirst()
                .orElse(null)

        if (sr != null) {
            return repository.read(repositoryId, sr.nodeId.id)
        }

        return null

        //        TreeNode result = base.children.stream().filter { it -> return it.name.toLowerCase() == name.toLowerCase() && it.type == NodeType.FOLDER }.findFirst().orElse(null)
        //        if (result == null && (level > 0 || level == -1)) {
        //            for (TreeNode child : base.children) {
        //                result = findFolder(child, name, level - 1)
        //
        //                if (result) {
        //                    return result
        //                }
        //            }
        //        }
        //        return result
    }

    private static void insertDepartmentList(XWPFParagraph paragraph, DepartmentNode node, int level) {
        if (!node.name.equals("Root")) {
            XWPFRun run = paragraph.createRun()

            // Добавляем отступы в зависимости от уровня
            for (int i = 0; i < level; i++) {
                run.addTab() // Добавляем табуляцию для отступа
            }

            run.setText(BLT + node.name) // Добавляем точку перед подразделением
            run.addBreak() // Переход на новую строку для следующего подразделения
        }

        // Рекурсивно добавляем дочерние подразделения
        for (DepartmentNode child : node.children) {
            insertDepartmentList(paragraph, child, level + 1)
        }
    }

    static FullModelDefinition getModelDefinition(ModelApi modelApi, String repositoryId, String modelId) {
        String key = "${repositoryId}-${modelId}".toString()

        if (cacheFullModelDefinition.containsKey(key)) {
            return cacheFullModelDefinition.get(key)
        } else {
            def value = modelApi.getModelDefinition(repositoryId, modelId, Arrays.asList("images", "modelType", "modelSymbols", "entries", "objectModelConnections", "parentObjectsInfo"))
            cacheFullModelDefinition.put(key, value)
            return value
        }

        /*if (modelId != FIRST_LEVEL_ID && result.model.parentNodeId.id.toLowerCase()){
            return null
        }
        return result
        */
    }

    static ObjectDefinitionNode getObjectDefinition(ObjectsApi objectsApi, String repositoryId, String objectId) {
        // Формируем ключ для кэша
        String key = "${repositoryId}-${objectId}".toString()

        // Проверяем кэш и получаем модель, если она уже там есть, иначе загружаем из modelApi
        if (cacheObjectDefinition.containsKey(key)) {
            return cacheObjectDefinition.get(key)
        } else {
            def value = objectsApi.getObjectDefinition(repositoryId, objectId)
            cacheObjectDefinition.put(key, value)
            return value
        }
    }

    // Метод для сбора всех путей от корневой организационной единицы до должности
    void collectOrgUnitPaths(OrgUnitNode node, List<String> currentPath, List<List<String>> paths) {
        // Добавляем текущий узел в путь
        currentPath.add(node.name)

        // Если у узла нет детей, это конец пути
        if (node.children.isEmpty()) {
            paths.add(new ArrayList<>(currentPath))
        } else {
            // Рекурсивно собираем пути для каждого дочернего узла
            for (OrgUnitNode child : node.children) {
                collectOrgUnitPaths(child, new ArrayList<>(currentPath), paths)
            }
        }
    }

    void traverseOrgUnitTree(OrgUnitNode node, int level, List<String> result) {
        // Формируем строку с отступами в зависимости от уровня
        String indent = BLT * level // Используем пробелы для отступа
        if (!node.isPosition) {
            result.add(indent + node.name)
        }

        List<OrgUnitNode> sortedChildren = node.children.sort().toList()
        // Рекурсивно добавляем всех потомков
        for (OrgUnitNode child : sortedChildren) {
            traverseOrgUnitTree(child, level + 1, result)
        }
    }

    // Класс OrgUnitNode с поддержкой множества бизнес-ролей
    static class OrgUnitNode implements Comparable<OrgUnitNode> {
        String id
        String repositoryId
        String name
        Set<OrgUnitNode> children = new TreeSet<>()
        Set<String> businessRoles = new TreeSet<>() // сортировка по алфавиту
        boolean isPosition = false

        OrgUnitNode(String id, String repositoryId, String name, boolean isPosition = false) {
            this.id = id
            this.repositoryId = repositoryId
            this.name = name
            this.isPosition = isPosition
        }

        OrgUnitNode(ObjectDefinitionNode node, boolean isPosition = false) {
            this.id = node.nodeId.id
            this.repositoryId = node.nodeId.repositoryId
            this.name = getNodeFullName(node)
            this.isPosition = isPosition
        }

        OrgUnitNode addChild(OrgUnitNode child) {
            children.add(child)
            return this
        }

        // Поиск дочернего узла по ID
        OrgUnitNode findChildById(String childId) {
            return children.find { it.id == childId }
        }

        boolean containsInChild(String childId) {
            if (this.id == childId) {
                return true
            }

            for (OrgUnitNode child : children) {
                if (child.containsInChild(childId)) {
                    return true
                }
            }

            return false
        }

        OrgUnitNode getNode(OrgUnitNode node) {
            if (this.id == node.id) {
                return this
            }

            for (OrgUnitNode child : children) {
                OrgUnitNode found = child.getNode(node)
                if (found != null) {
                    return found
                }
            }

            return null
        }

        boolean equals(o) {
            if (this.is(o)) return true
            if (o == null || getClass() != o.class) return false

            OrgUnitNode that = (OrgUnitNode) o

            if (id != that.id) return false

            return true
        }

        int hashCode() {
            return (id != null ? id.hashCode() : 0)
        }

        /**
         * OrgUnitNode содержит должность или название орг. единицы
         * В таблице 1 сначала выводим должности, потом орг. единицы
         */
        @Override
        int compareTo(OrgUnitNode o2) {
            if (this.isPosition && o2.isPosition) {
                if (this.name == o2.name) {
                    return this.id <=> o2.id
                }

                return this.name <=> o2.name
            } else if (this.isPosition && !o2.isPosition) {
                return -1
            } else if (!this.isPosition && o2.isPosition) {
                return 1
            } else {
                return unitFlatComparator.compare(this.name, o2.name)
            }
        }

        void printTree(Integer level = 1) {
            System.out.println("  " * level + this.id + " : " + this.name)
            for (OrgUnitNode child : this.children) {
                child.printTree(level + 1)
            }
        }
    }

    /**
     * Построение последовательности родительских узлов для должности и добавление полученной ветки в дерево
     * @param forest корневые узлы деревьев
     * @param positionNode - должность
     * @param objectsApi - АПИ для запроса определения объекта
     */
    void addPositionToTree(Map<String, OrgUnitNode> forest, OrgUnitNode positionNode, ObjectsApi objectsApi) {
        // включает CT_IS_CRT_BY
        OrgUnitNode head = positionNode
        OrgUnitNode child = positionNode
        OrgUnitNode parentUnit = getParentUnit(child, objectsApi)

        // является линейным руководителем для CT_IS_SUPERIOR_1
        while (parentUnit != null) {
            for (OrgUnitNode root : forest.values()) {
                OrgUnitNode node = root.getNode(parentUnit)
                if (node != null) {
                    node.addChild(child)
                    return
                }
            }

            head = parentUnit
            parentUnit.addChild(child)

            child = parentUnit
            parentUnit = getParentUnit(parentUnit, objectsApi)
        }

        forest.put(head.id, head)
    }


    /**
     * Определение родительского узла через связи на моделях
     * @param childUnit узел, для которого ищем родителя
     * @param objectsApi АПИ для запроса определения объекта
     * @return родительский узел
     */
    OrgUnitNode getParentUnit(OrgUnitNode childUnit, ObjectsApi objectsApi) {
        if (FAST_ORG_STRUCT) {
            return getParentUnitFast(childUnit, objectsApi)
        }

        String relationTypeId = childUnit.isPosition ? "CT_IS_CRT_BY" : "CT_IS_SUPERIOR_1"
        ObjectDefinitionNode node = getObjectDefinition(objectsApi, childUnit.repositoryId, childUnit.id)

        List<ObjectConnection> omc = node.getObjectModelConnections()
                .stream()
                .flatMap { it.getConnections().stream() }
                .toList()

        for (ObjectConnection connection : omc) {
            if (connection.connectedObjectTypeId == "OT_ORG_UNIT"
                    && connection.edgeTypeId.equals(relationTypeId)
                    && connection.isOutgoingEdge == false) {

                ObjectDefinitionNode parentNodeDefinition = getObjectDefinition(objectsApi, node.getNodeId().getRepositoryId(), connection.getConnectedObjectDefinitionId())
                return new OrgUnitNode(parentNodeDefinition)
            }
        }

        return null
    }

    /**
     * Определение родительского узла через определения связей в орг. структуре
     * @param childUnit узел, для которого ищем родителя
     * @return родительский узел
     */
    OrgUnitNode getParentUnitFast(OrgUnitNode childUnit, ObjectsApi objectsApi) {
        String relationTypeId = childUnit.isPosition ? "CT_IS_CRT_BY" : "CT_IS_SUPERIOR_1"

        def targetRelations = relationsTargetMap.get(childUnit.id)

        if (targetRelations != null) {
            def parentId = targetRelations.stream()
                    .filter { (it.edgeTypeId == relationTypeId) }
                    .map { it.sourceObjectDefinitionId }
                    .findFirst()
                    .orElse(null)

            if (parentId != null) {
                ObjectDefinitionNode fastUnit = objectMap.get(parentId)
                ObjectDefinitionNode unit = getObjectDefinition(objectsApi, fastUnit.getNodeId().getRepositoryId(), fastUnit.getNodeId().getId())
                return new OrgUnitNode(unit)
            }
        }

        return null
    }

    static String getNodeFullName(ObjectDefinitionNode node) {
        try {
            String name = node.getName()

            AttributeValue attributeValue = node.getAttributes().stream()
                    .filter { it.typeId == "AT_NAME_FULL" }
                    .findFirst()
                    .orElse(null)

            if (attributeValue != null
                    && attributeValue.value != null
                    && !attributeValue.value.trim().isEmpty()) {
                name = attributeValue.value
            }

            if (name != null) {
                trimName(name)
            } else {
                return "[название отсутствует]"
            }
        } catch (Exception e) {
            log.error(e.toString())
        }
    }

    static String trimName(String name) {
        return name.replaceAll("[\\s\\n]+", " ").trim()
    }

    private static void saveImageTable(String path, List<ProcessPng> table) {
        Path filePath = Paths.get(path)

        try {
            Files.deleteIfExists(filePath);
            try (ObjectOutputStream oos = new ObjectOutputStream(new FileOutputStream(filePath.toFile()))) {
                oos.writeObject(table);
            }
        } catch (Exception e) {
            log.error(e.toString())
        }
    }

    class DocumentCreator {
        FileNodeDTO createDoc(String templateName,
                              String fileName,
                              DocumentData data) {
            XWPFDocument doc = getTemplate(templateName)

            doc = fillDoc(doc, data)

            byte[] bytes = new byte[0]
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                doc.write(outputStream)
                bytes = outputStream.toByteArray()
            }

            def fileRepositoryId = 'file-folder-root-id'
            def userId = context.principalId()

            def fileNode = FileNodeDTO.builder()
                    .nodeId(NodeId.builder().id(UUID.randomUUID().toString()).repositoryId(fileRepositoryId).build())
                    .parentNodeId(NodeId.builder().id(String.valueOf(userId)).repositoryId(fileRepositoryId).build())
                    .extension('docx')
                    .file(new SimpleMultipartFile(fileName, bytes))
                    .name(fileName + ".docx")
                    .build()

            if (!DEBUG_LOCAL) {
                context.getApi(FileApi).uploadFile(fileNode)
            }
            return fileNode
        }

        XWPFDocument getTemplate(String templateName) {
            if (DEBUG_LOCAL) {
                // Указываем путь к файлуC:\Users\bovae\IdeaProjects\sila\templates\reg_bp.docx
                String filePath = "${LOCAL_PATH}\\templates\\${templateName}"

                // Создаем объект File
                File file = new File(filePath)

                // Проверяем, существует ли файл
                if (!file.exists()) {
                    throw new IOException("Файл не найден: " + filePath)
                }

                // Читаем файл в массив байтов
                try {
                    FileInputStream fis = new FileInputStream(file)
                    return new XWPFDocument(fis)
                } catch (Exception e) {
                    log.error("Ошибка чтения файла", e)
                }
            }
            //   Определение шаблона
            TreeNode fileTreeNode = null
            def fileFolderTreeNode = context.createTreeRepository(false).read(TEMPLATE_FOLDER_ID, TEMPLATE_FOLDER_ID)
            def childrens = fileFolderTreeNode.getChildren()
            for (TreeNode children : childrens) {
                if (children.getType().name().equals("FILE_FOLDER") && children.getName().equals("Общие")) {
                    def files = children.getChildren()
                    for (TreeNode file : files) {
                        if (file.getName().toLowerCase().equals(templateName.toLowerCase())) {
                            fileTreeNode = file
                            break
                        }
                    }
                    if (fileTreeNode != null) {
                        break
                    }
                }
            }
            //  TODO   Надо добавить вывод сообщения, что не найден шаблон
            byte[] file = context.getApi(FileApi.class).downloadFile(TEMPLATE_FOLDER_ID, fileTreeNode.id)
            return new XWPFDocument(new ByteArrayInputStream(file))
        }

        static XWPFDocument fillDoc(XWPFDocument doc, DocumentData data) {
            AtomicInteger appCount = new AtomicInteger(1) // mutable - чтобы изменять внутри метода
            ZipSecureFile.setMinInflateRatio(0)

            documentNumbering = createNumberingStyle(doc)

            processSimpleReplace(doc, data.parameters)

            findParagraphByText(doc, "<Приложение схема подпроцесса>")
                    .ifPresent { p -> processDiagram(p, data.subProcesses, appCount) }

            findParagraphByText(doc, "<Приложение схема декомпозиции подпроцесса>")
                    .ifPresent { p -> processDiagram(p, data.decSubProcesses, appCount) }

            findParagraphByText(doc, "<Пункт подпроцесс роли>")
                    .ifPresent { p -> processRole(p, data.createTablesToData) }

            processTables(doc, data.numberTablesToData)

            for (String parameter : data.parameterLists.keySet()) {
                findParagraphByText(doc, "<" + parameter + ">")
                        .ifPresent { p -> processParameterList(p, parameter, data.parameterLists.get(parameter)) }
            }

            findParagraphByText(doc, "ЛИСТ СОГЛАСОВАНИЯ")
                    .ifPresent { p -> processAcceptList(p, data.orgUnitNames) }

            List<XWPFParagraph> records = new ArrayList<XWPFParagraph>()
            for (XWPFParagraph p : doc.getParagraphs()) {
                try {
                    if (p.getParagraphText().contains("todelete")) {
                        records.add(p)
                    }
                } catch (Exception e) {
                    System.err.println(e.getMessage())
                }
            }
            for (int i = 0; i < records.size(); i++) {
                doc.removeBodyElement(doc.getPosOfParagraph(records.get(i)))
            }

            // #76366 Обновление содержания
            doc.enforceUpdateFields()

            return doc
        }

        static BigInteger createNumberingStyle(XWPFDocument doc) {
            // Добавляем стиль нумерации в документ
            XWPFNumbering numbering = doc.getNumbering();
            if (numbering == null) {
                numbering = doc.createNumbering();
            }

            def abstractNumId = numbering.getAbstractNums().stream()
                    .map { it.getAbstractNum().getAbstractNumId() }
                    .max { o1, o2 -> o1.longValue() <=> o2.longValue() }
                    .orElse(null);

            // Создаём новый стиль нумерации
            CTAbstractNum abstractNum = CTAbstractNum.Factory.newInstance();
            abstractNum.setAbstractNumId(abstractNumId.add(BigInteger.ONE)); // Избегаем 0

            for (int level = 0; level < 5; level++) {
                CTLvl cTLvl = abstractNum.addNewLvl()
                cTLvl.setIlvl(BigInteger.valueOf(level))
                cTLvl.addNewNumFmt().setVal(STNumberFormat.BULLET)
                cTLvl.addNewLvlText().setVal(bulletSymbol.toString())
                cTLvl.addNewLvlJc().setVal(STJc.LEFT)
                cTLvl.addNewRPr()
                CTFonts f = cTLvl.getRPr().addNewRFonts()
                f.setAscii(bulletFont)
                f.setHAnsi(bulletFont)
            }

            XWPFAbstractNum xwpfAbstractNum = new XWPFAbstractNum(abstractNum);
            BigInteger abstractNumID = numbering.addAbstractNum(xwpfAbstractNum);

            // Создаём экземпляр нумерации
            BigInteger numID = numbering.addNum(abstractNumID);
            return numID
        }


        static void createTable(XWPFDocument doc, XWPFParagraph positionParagraph, int cols, List<List<String>> data) {
            // Создаем таблицу с одной строкой и нужным количеством столбцов
            XmlCursor cursor = positionParagraph.getCTP().newCursor()
            XWPFTable table = doc.insertNewTbl(cursor)
            for (int i = 0; i < cols - 1; i++) {
                table.getRow(0).createCell()
            }

            if (table == null) {
                return
            }

            // Удаляем первую строку, если нужно добавить данные с нуля
            if (data == null) {
                table.removeRow(0)
                return
            }

            // Настройка границ таблицы
            table.setInsideHBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.setInsideVBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(6000))
            table.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2000))

            // Устанавливаем текст первой строки, которая уже создана
            XWPFTableRow headerRow = table.getRow(0)
            for (int i = 0; i < cols; i++) {
                XWPFTableCell cell = headerRow.getCell(i)
                XWPFParagraph paragraph = cell.getParagraphArray(0)
                paragraph.setSpacingBefore(Math.round(11.3 * 20).intValue())
                paragraph.setSpacingAfter(Math.round(11.3 * 20).intValue())

                if (i < data.get(0).size()) {
                    String trimmedValue = data.get(0).get(i).replaceAll("\\s+", " ").trim()
                    addText(paragraph, trimmedValue, 12)
                }
            }

            headerRow.setRepeatHeader(true)

            // Добавляем остальные строки в таблицу
            for (int r = 1; r < data.size(); r++) {
                List<String> rowData = data.get(r)
                addRow(table, rowData)
            }
        }

        static void updateTable(XWPFDocument doc, int tableIndex, List<List<String>> data) {
            XWPFTable table = doc.getTableArray(tableIndex)

            if (table == null) {
                return
            }

            if (table.getNumberOfRows() > 1) {
                table.removeRow(1)
            }

            if (data == null) {
                doc.removeBodyElement(doc.getPosOfTable(table))

                return
            }

            // Настройка границ таблицы
            table.setInsideHBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.setInsideVBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 1, 0, "000000")
            table.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(6000))
            table.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2000))

            // Добавляем строки в таблицу
            for (List<String> rowData : data) {
                addRow(table, rowData)
            }
        }


        static void addRow(XWPFTable table, List<String> rowData) {
            XWPFTableRow newRow = table.createRow()

            for (int i = 0; i < rowData.size(); i++) {
                XWPFTableCell cell = newRow.getCell(i)
                if (cell == null) {
                    cell = newRow.createCell()
                }
                cell.removeParagraph(0)

                String[] texts = rowData.get(i).split('/n')
                if (texts.size() > 1) {
                    String listId = UUID.randomUUID().toString()
                    texts.each {
                        XWPFParagraph paragraph = cell.addParagraph()
                        paragraph.setSpacingBefore(Math.round(2.8 * 20).intValue())
                        paragraph.setSpacingAfter(Math.round(2.8 * 20).intValue())

                        String trimmedValue = it.replaceAll("\\s+", " ").trim()
                        addText(paragraph, trimmedValue, 12, listId)
//                            paragraph.createRun()
                    }
                } else {
                    XWPFParagraph paragraph = cell.addParagraph()
                    paragraph.setSpacingBefore(Math.round(2.8 * 20).intValue())
                    paragraph.setSpacingAfter(Math.round(2.8 * 20).intValue())

                    String trimmedValue = rowData.get(i).replaceAll("\\s+", " ").trim()
                    addText(paragraph, trimmedValue, 12)
                }


                // Настройка границ ячеек
                cell.getCTTc().addNewTcPr().addNewTcBorders().addNewTop().setVal(STBorder.SINGLE)
                cell.getCTTc().addNewTcPr().addNewTcBorders().addNewBottom().setVal(STBorder.SINGLE)
                cell.getCTTc().addNewTcPr().addNewTcBorders().addNewLeft().setVal(STBorder.SINGLE)
                cell.getCTTc().addNewTcPr().addNewTcBorders().addNewRight().setVal(STBorder.SINGLE)
            }

        }

        static void processTables(XWPFDocument doc, Map<Integer, List<List<String>>> numberTablesToData) {
            int countDeletedTable = 0
            for (Map.Entry<Integer, List<List<String>>> table : numberTablesToData) {
                if (table != null) {
                    int numberTable = Integer.valueOf(table.getKey()) - countDeletedTable
                    updateTable(doc, numberTable, table.getValue())

                    if (table.getValue() == null) {
                        countDeletedTable++
                    }
                }
            }
        }

        static void processSimpleReplace(XWPFDocument doc, Map<String, String> parameters) {
            for (String parameter : parameters.keySet()) {
                String placeholder = "<" + parameter + ">"
                for (XWPFParagraph paragraph : doc.getParagraphs()) {
                    replaceText(paragraph, placeholder, parameters.get(parameter))
                }

                //замена текста в колонтитулах
                for (XWPFHeader header : doc.getHeaderList()) {
                    for (XWPFParagraph headerParagraph : header.getParagraphs()) {
                        replaceText(headerParagraph, placeholder, parameters.get(parameter))
                    }
                }

                for (XWPFTable table : doc.getTables()) {
                    for (XWPFTableRow row : table.getRows()) {
                        for (XWPFTableCell cell : row.getTableCells()) {
                            for (XWPFParagraph cellParagraph : cell.getParagraphs()) {
                                replaceText(cellParagraph, placeholder, parameters.get(parameter));
                            }
                        }
                    }
                }
            }
        }

        static void processRole(XWPFParagraph p, Map<String, List<List<String>>> createTablesToData) {
            XWPFDocument doc = p.getDocument()

            int count = 1
            for (Map.Entry<String, List<List<String>>> table : createTablesToData) {
                List<List<String>> tableValue = table.getValue()
                if (tableValue.size() > 1) {
                    String name = table.getKey().replaceAll("\\s+", " ").trim()

                    XWPFParagraph titleParagraph = insertParagraph(doc, p)
                    addText(titleParagraph, new StringBuilder("\\B4.${count}. ${name}").toString(), null)
                    titleParagraph.setStyle("RegHeaderStyle2")

                    createTable(doc, p, 7, tableValue)
                    count++
                }
            }
            removeParagraph(doc, p)
        }

        static void processDiagram(XWPFParagraph p, List<ProcessPng> subProcesses, AtomicInteger appCount) {
            XWPFDocument doc = p.getDocument()

            for (int i = 0; i < subProcesses.size(); i++) {
                ProcessPng subProcess = subProcesses.get(i)
                String caption = "\\BПРИЛОЖЕНИЕ ${appCount.get()}. МОДЕЛЬ БИЗНЕС-ПРОЦЕССА ${subProcess.name}".toString()

                // Создаём новый параграф для заголовка приложения
                XWPFParagraph titleParagraph = insertParagraph(doc, p, p)
                addText(titleParagraph, caption, null)
                titleParagraph.setStyle("RegHeaderStyle")

                // Создаём параграф для изображения
                XWPFParagraph imageParagraph = insertParagraph(doc, p)
                insertPicture(imageParagraph, subProcess.getPng(), caption.length())

                appCount.getAndIncrement()
            }

            removeParagraph(doc, p)
        }

        static void processParameterList(XWPFParagraph p, String parameter, List<String> parameterValues) {
            XWPFDocument doc = p.getDocument()
            String uuid = UUID.randomUUID().toString()

            for (String s : parameterValues) {
                if ((parameter.equals("состав_групп") || parameter.equals("состав_документов")) && s.startsWith('\\B')) {
                    XWPFParagraph newParagraph = insertParagraph(doc, p)
                    addText(newParagraph, s, null, uuid)
                    newParagraph.setStyle("RegHeaderStyle2")
                } else {
                    XWPFParagraph newParagraph = insertParagraph(doc, p, p)
                    addText(newParagraph, s, null, uuid)
                }
            }
            removeParagraph(doc, p)
        }

        static void processAcceptList(XWPFParagraph p, List<String> orgUnitNames) {
            XWPFDocument doc = p.getDocument()

            orgUnitNames = orgUnitNames.stream().sorted(unitTreeComparator)
                    .filter { it != null && !it.isEmpty() }
                    .toList()

            if (orgUnitNames.size() == 1) {
                findParagraphByText(doc, "<Лист_согласования>")
                        .ifPresent { pp -> replaceText(pp, "<Лист_согласования>", orgUnitNames.get(0)) }

            } else {
                int parPos = doc.getPosOfParagraph(p)

                List<IBodyElement> template = new ArrayList<>()
                for (int i = parPos; i < doc.getBodyElements().size(); i++) {
                    IBodyElement bodyElement = doc.getBodyElements().get(i)
                    template.add(bodyElement)
                    parPos++

                    if (bodyElement instanceof XWPFTable) {
                        break
                    }
                }

                XWPFParagraph positionParagraph = doc.getBodyElements().get(parPos) as XWPFParagraph

                try {
                    for (int i = 1; i < orgUnitNames.size(); i++) {
                        XWPFParagraph breakPar = doc.insertNewParagraph(positionParagraph.getCTP().newCursor())
                        breakPar.createRun().addBreak(BreakType.PAGE)

                        for (IBodyElement bodyElement : template) {
                            if (bodyElement instanceof XWPFParagraph) {
                                XWPFParagraph newParagraph = insertParagraph(doc, positionParagraph, bodyElement as XWPFParagraph)
                            } else if (bodyElement instanceof XWPFTable) {
                                XWPFTable newTable = insertTable(doc, positionParagraph, bodyElement as XWPFTable)
                            }
                        }
                    }

                    for (def orgUnitName : orgUnitNames) {
                        findParagraphByText(doc, "<Лист_согласования>")
                                .ifPresent { pp -> replaceText(pp, "<Лист_согласования>", orgUnitName) }
                    }
                } catch (Exception e) {
                    e.printStackTrace()
                }
            }
        }

        static void removeParagraph(XWPFDocument doc, XWPFParagraph paragraph) {
            def pos = doc.getPosOfParagraph(paragraph)
            doc.removeBodyElement(pos)
        }


        static void cloneStyle(XWPFParagraph target, XWPFParagraph source) {
            CTPPr pPr = target.getCTP().isSetPPr() ? target.getCTP().getPPr() : target.getCTP().addNewPPr()
            pPr.set(source.getCTP().getPPr())

            //            if (!source.getRuns().isEmpty()) {
            //                for (XWPFRun targetRun : target.getRuns()) {
            //                    CTRPr rPr = targetRun.getCTR().isSetRPr() ? targetRun.getCTR().getRPr() : targetRun.getCTR().addNewRPr()
            //                    rPr.set(source.getRuns().get(0).getCTR().getRPr())
            //                }
            //            }

            //            for (XWPFRun r : source.getRuns()) {
            //                XWPFRun nr = clone.createRun();
            //                cloneRun(nr, r);
            //            }
        }


        static void cloneParagraph(XWPFParagraph source, XWPFParagraph target) {
            CTPPr pPr = target.getCTP().isSetPPr() ? target.getCTP().getPPr() : target.getCTP().addNewPPr();
            pPr.set(source.getCTP().getPPr());
            for (XWPFRun r : source.getRuns()) {
                XWPFRun nr = target.createRun();
                cloneRun(nr, r);
            }
        }

        static void cloneRun(XWPFRun target, XWPFRun source) {
            CTRPr rPr = target.getCTR().isSetRPr() ? target.getCTR().getRPr() : target.getCTR().addNewRPr();
            rPr.set(source.getCTR().getRPr());
            target.setText(source.getText(0));
        }

        static XWPFParagraph insertParagraph(XWPFDocument doc, XWPFParagraph positionParagraph) {
            XmlCursor cursor = positionParagraph.getCTP().newCursor()
            XWPFParagraph emptyParagraph = doc.insertNewParagraph(cursor)
            return emptyParagraph
        }

        static XWPFParagraph insertParagraph(XWPFDocument doc, XWPFParagraph positionParagraph, XWPFParagraph paragraph) {
            XmlCursor cursor = positionParagraph.getCTP().newCursor()
            XWPFParagraph emptyParagraph = doc.insertNewParagraph(cursor)
            cloneParagraph(paragraph, emptyParagraph)
            //            emptyParagraph.getCTP().set(paragraph.getCTP().copy())
            return emptyParagraph
        }

        static XWPFTable insertTable(XWPFDocument doc, XWPFParagraph positionParagraph, XWPFTable table) {
            XmlCursor cursor2 = positionParagraph.getCTP().newCursor()
            XWPFTable newTable = doc.insertNewTbl(cursor2)
            cloneTable(table, newTable)
            //            XWPFTable tableCopy = new XWPFTable((CTTbl)table.getCTTbl().copy(), doc);
            //            doc.setTable(tablePosition, tableCopy)
            return newTable
        }

        static void cloneTable(XWPFTable source, XWPFTable target) {
            target.getCTTbl().setTblPr(source.getCTTbl().getTblPr());
            target.getCTTbl().setTblGrid(source.getCTTbl().getTblGrid());
            for (int r = 0; r < source.getRows().size(); r++) {
                XWPFTableRow targetRow = target.createRow();
                XWPFTableRow row = source.getRows().get(r);
                targetRow.getCtRow().setTrPr(row.getCtRow().getTrPr());
                for (int c = 0; c < row.getTableCells().size(); c++) {
                    //newly created row has 1 cell
                    XWPFTableCell targetCell = c == 0 ? targetRow.getTableCells().get(0) : targetRow.createCell();
                    XWPFTableCell cell = row.getTableCells().get(c);
                    targetCell.getCTTc().setTcPr(cell.getCTTc().getTcPr());
                    XmlCursor cursor = targetCell.getParagraphArray(0).getCTP().newCursor();
                    for (int p = 0; p < cell.getBodyElements().size(); p++) {
                        IBodyElement elem = cell.getBodyElements().get(p);
                        if (elem instanceof XWPFParagraph) {
                            XWPFParagraph targetPar = targetCell.insertNewParagraph(cursor);
                            cursor.toNextToken();
                            XWPFParagraph par = (XWPFParagraph) elem;
                            cloneParagraph(par, targetPar);
                        } else if (elem instanceof XWPFTable) {
                            XWPFTable targetTable = targetCell.insertNewTbl(cursor);
                            XWPFTable table = (XWPFTable) elem;
                            copyTable(table, targetTable);
                            cursor.toNextToken();
                        }
                    }
                    //newly created cell has one default paragraph we need to remove
                    targetCell.removeParagraph(targetCell.getParagraphs().size() - 1);
                }
            }
            //newly created table has one row by default. we need to remove the default row.
            target.removeRow(0);
        }

        static int getParagraphPosition(XWPFDocument doc, XWPFParagraph paragraph) {
            for (int i = 0; i < doc.getParagraphs().size(); i++) {
                if (doc.getParagraphs().get(i) == paragraph) {
                    return i
                }
            }

            int bodyPosition = doc.getPosOfParagraph(paragraph)
            return doc.getParagraphPos(bodyPosition)
        }

        static XWPFParagraph insertPicture(XWPFParagraph imageParagraph, byte[] image, int captionLength) {
            try {
                if (image.length == 0) {
                    throw new Exception("Изображение не найдено")
                }

                XWPFRun run = imageParagraph.createRun()
                imageParagraph.setAlignment(ParagraphAlignment.CENTER)

                final int longSide = 700
                final int shortSide = 450
                final int captionStringHeight = 16

                int pageW
                int pageH
                int width
                int height

                try (InputStream is = new ByteArrayInputStream(image)) {
                    BufferedImage img = ImageIO.read(is)

                    // Если изображение широкое, то разворачиваем страницу
                    if (img.width > img.height) {
                        setPageOrientation(imageParagraph, STPageOrientation.LANDSCAPE)
                        int captionStringCount = (int) Math.ceil(captionLength / 95.0) // 95 букв в одной строке
                        pageW = longSide
                        pageH = shortSide - (captionStringCount + 1) * captionStringHeight
                        // + 1 потому что первая строка очень высокая
                    } else {
                        setPageOrientation(imageParagraph, STPageOrientation.PORTRAIT)
                        int captionStringCount = (int) Math.ceil(captionLength / 57.0) // 57 букв в одной строке
                        pageW = shortSide
                        pageH = longSide - (captionStringCount + 1) * captionStringHeight
                        // + 1 потому что первая строка очень высокая
                    }

                    // Если изображение не помещается на страницу, то надо его масштабировать
                    double scale = 1.0
                    if (img.width > pageW || img.height > pageH) {
                        double widthScale = pageW / img.getWidth()
                        double heightScale = pageH / img.getHeight()
                        scale = Math.min(widthScale, heightScale)
                    }

                    width = (int) (img.getWidth() * scale)
                    height = (int) (img.getHeight() * scale)
                }

                try (InputStream is = new ByteArrayInputStream(image)) {
                    run.addPicture(is, XWPFDocument.PICTURE_TYPE_PNG, "image.png", Units.toEMU(width), Units.toEMU(height))
                }
            } catch (Exception e) {
                setPageOrientation(imageParagraph, STPageOrientation.PORTRAIT)
                addText(imageParagraph, "Ошибка вставки изображения: ${e.getMessage()}", null)
            }

            return imageParagraph
        }

        static void setPageOrientation(XWPFParagraph paragraph, STPageOrientation.Enum orientation) {
            def sect = paragraph.getCTPPr().addNewSectPr()
            def pageSize = sect.addNewPgSz()
            pageSize.setOrient(orientation)
            if (orientation == STPageOrientation.LANDSCAPE) {
                pageSize.setW(842 * 20)
                pageSize.setH(595 * 20)
            } else {
                pageSize.setW(595 * 20)
                pageSize.setH(842 * 20)
            }
        }


        static Stream<XWPFParagraph> findParagraphStreamByText(IBody body, String text) {
            return body.getParagraphs()
                    .stream()
                    .filter { it.getText().contains(text) }
        }

        static Optional<XWPFParagraph> findParagraphByText(IBody body, String text) {
            return findParagraphStreamByText(body, text).findFirst()
        }

        static void replaceText(XWPFParagraph paragraph, String target, String replacement) {
            if (paragraph.getText().contains(target)) {
                def newText = paragraph.getText().replace(target, replacement)
                addText(paragraph, newText, null)
            }
        }

        static void addText(XWPFParagraph paragraph, String text, Integer fontSize, String listId = null) {
            XWPFRun run;
            if (CollectionUtils.isNotEmpty(paragraph.getRuns())) {
                // Оставляем только первый run
                while (paragraph.getRuns().size() > 1) {
                    paragraph.removeRun(1);
                }
                run = paragraph.getRuns().get(0);
            } else {
                run = paragraph.createRun();
            }

            // Заменяем двойные кавычки на одинарные, ибо таковы правила русского языка
            if (text.contains('""')) {
                text = text.replaceAll('""', '"')
            }
            // Определяем уровень вложенности по количеству символов BLT
            int bulletCount = StringUtils.countMatches(text, BLT)
            text = text.replaceAll(BLT, '');

            if (bulletCount > 0) {
                // Получаем или создаём ID нумерации для текущего уровня
                //                BigInteger numID = getBulletNumberingId(paragraph.getDocument(), bulletCount - 1, listId);
                paragraph.setNumID(documentNumbering)
                paragraph.setNumILvl(BigInteger.valueOf(bulletCount - 1))
                paragraph.setIndentationLeft(CM1 + bulletCount * CM05); // Задаём отступы в зависимости от уровня
                paragraph.setIndentationHanging(CM05); // Задаём отступы в зависимости от уровня
            }

            // Обработка стилей, начинающихся с "\"
            while (text.startsWith("\\")) {
                if (text.startsWith("\\B")) {
                    text = text.substring(2);
                    run.setBold(true);
                } else if (text.startsWith("\\!B")) {
                    text = text.substring(3);
                    run.setBold(false);
                } else if (text.startsWith("\\I")) {
                    text = text.substring(2);
                    run.setItalic(true);
                } else if (text.startsWith("\\U")) {
                    text = text.substring(2);
                    run.setUnderline(UnderlinePatterns.SINGLE);
                } else if (text.startsWith("\\L")) {
                    text = text.substring(2);
                    try {
                        String level = text.substring(0, 1)
                        paragraph.setIndentationLeft(level.toInteger() * CM1); // Задаём отступы в зависимости от уровня
                        text = text.substring(1)
                    } catch (Exception e) {
                        paragraph.setIndentationLeft(CM1)
                    }
                } else if (text.startsWith("\\A")) {
                    paragraph.setIndentationFirstLine(CM1)
                    paragraph.setIndentationLeft(0)
                    paragraph.setIndentationRight(0)
                    paragraph.setSpacingBefore(Math.round(5.65 * 20).intValue())
                    paragraph.setSpacingAfter(0)
                    paragraph.setAlignment(ParagraphAlignment.BOTH)
                    text = text.substring(2)
                } else {
                    text = text.substring(1);
                }
            }

            if (fontSize != null && fontSize > 0) {
                run.setFontSize(fontSize);
            }

            def pos = text.indexOf("<>")
            if (pos >= 0) {
                def text1 = text.substring(0, pos)
                run.setText(text1, 0)
                paragraph.addRun(run)

                XWPFRun clone1 = paragraph.createRun()
                CTRPr rPr1 = clone1.getCTR().isSetRPr() ? clone1.getCTR().getRPr() : clone1.getCTR().addNewRPr()
                rPr1.set(run.getCTR().getRPr())
                clone1.setText("<>")
                clone1.getCTR().addNewRPr().addNewHighlight().setVal(STHighlightColor.YELLOW)
                paragraph.addRun(clone1)

                if (pos + 2 < text.length()) {
                    def text3 = text.substring(pos + 2, text.length())
                    XWPFRun clone2 = paragraph.createRun()
                    CTRPr rPr2 = clone2.getCTR().isSetRPr() ? clone2.getCTR().getRPr() : clone2.getCTR().addNewRPr()
                    rPr2.set(run.getCTR().getRPr())
                    clone2.setText(text3)
                    paragraph.addRun(clone2)
                }
            } else {
                run.setText(text, 0)
                paragraph.addRun(run)
            }
        }

        /**
         * Получает или создаёт ID нумерации для заданного уровня.
         *
         * @param doc Документ Word
         * @param level Уровень вложенности (0 для первого уровня, 1 для второго и т.д.)
         * @return ID нумерации
         */
        static BigInteger getBulletNumberingId(XWPFDocument doc, int level) {
            if (bulletNumberingMap.containsKey(level)) {
                return bulletNumberingMap.get(level);
            }

            // Создаём новый стиль нумерации
            CTAbstractNum abstractNum = CTAbstractNum.Factory.newInstance();
            abstractNum.setAbstractNumId(BigInteger.valueOf(level)); // Избегаем 0

            CTLvl cTLvl = abstractNum.addNewLvl();
            cTLvl.setIlvl(BigInteger.valueOf(level));
            cTLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
            cTLvl.addNewLvlText().setVal(bulletSymbol.toString());
            cTLvl.addNewLvlJc().setVal(STJc.LEFT);

            cTLvl.addNewRPr()
            CTFonts f = cTLvl.getRPr().addNewRFonts()
            f.setAscii("Symbol")
            f.setHAnsi("Symbol")

            XWPFAbstractNum xwpfAbstractNum = new XWPFAbstractNum(abstractNum);

            // Добавляем стиль нумерации в документ
            XWPFNumbering numbering = doc.getNumbering();
            if (numbering == null) {
                numbering = doc.createNumbering();
            }
            BigInteger abstractNumID = numbering.addAbstractNum(xwpfAbstractNum);

            // Создаём экземпляр нумерации
            BigInteger numID = numbering.addNum(abstractNumID);

            // Сохраняем ID в карту для переиспользования
            //            bulletNumberingMap.put(key, numID);
            return numID
        }
    }

    static class DocumentData implements Serializable {
        long serialVersionUID = 42L

        Map<String, String> parameters
        Map<String, List<String>> parameterLists
        Map<Integer, List<List<String>>> numberTablesToData
        Map<String, List<List<String>>> createTablesToData
        List<ProcessPng> subProcesses
        List<ProcessPng> decSubProcesses
        List<String> orgUnitNames

        DocumentData() {
        }
    }

    /**
     * Последовательное сравнение по весам подразделений, по номерам подразделений, по алфавиту
     */
    static class OrgUnitFlatComparator implements Comparator<String> {

        private int getOrgUnitWeight(String namePart) {
            switch (namePart) {
                case 'Департамент':
                    return 1
                case 'Управление':
                    return 2
                case 'Отдел':
                    return 3
                case 'отдел':
                    return 4
                case 'Секретариат':
                    return 4
                default:
                    return Integer.MAX_VALUE
            };
        }

        private int compareNameParts(String np1, String np2) {
            return getOrgUnitWeight(np1) - getOrgUnitWeight(np2);
        }

        private int compareNumberParts(String np1, String np2) {
            String[] numberParts1 = np1.replaceAll("\\s", "").split("/");
            String[] numberParts2 = np2.replaceAll("\\s", "").split("/");

            int maxLength = Math.max(numberParts1.length, numberParts2.length);
            int cr = 0;
            for (int i = 0; i < maxLength; i++) {
                if (numberParts1.length > i && numberParts2.length > i) {
                    cr = Integer.parseInt(numberParts1[i]) - Integer.parseInt(numberParts2[i]);
                } else if (numberParts1.length > i) {
                    cr = 1;
                } else {
                    cr = -1;
                }

                if (cr != 0) {
                    return cr;
                }
            }

            return cr;
        }

        private int compareNames(String name1, String name2) {
            try {
                String[] nameParts1 = StringUtils.split(name1, null, 2);
                String[] nameParts2 = StringUtils.split(name2, null, 2);

                int namePartCompareResult = compareNameParts(nameParts1[0], nameParts2[0]);
                if (namePartCompareResult != 0) {
                    return namePartCompareResult;
                } else {
                    if (nameParts1.length == 2 && nameParts2.length == 2) {
                        if (isNumber(nameParts1[1]) && isNumber(nameParts2[1])) {
                            return compareNumberParts(nameParts1[1], nameParts2[1]);
                        }
                    }
                    return name1.compareTo(name2);
                }
            } catch (Exception e) {
                log.warn("Ошибка сравнения названий орг. единиц: [{}] и [{}]", name1, name2)
                return 0;
            }
        }

        private boolean isNumber(String s) {
            String ss = s.replaceAll("\\s", "");
            return onlyNumberPattern.matcher(ss).matches();
        }

        @Override
        int compare(String name1, String name2) {
            return compareNames(name1, name2);
        }
    }

    /**
     * Сравнение по номерам в названии
     */
    static class OrgUnitTreeComparator implements Comparator<String> {
        @Override
        int compare(String o1, String o2) {
            try {
                String[] firstObject = o1.replaceAll("[^0-9/]", "").split("/")
                String[] secondObject = o2.replaceAll("[^0-9/]", "").split("/")

                int length = Math.min(firstObject.length, secondObject.length)

                for (int i = 0; i < length; i++) {
                    int num1 = Integer.parseInt((firstObject[i]))
                    int num2 = Integer.parseInt((secondObject[i]))

                    if (num1 != num2) {
                        return Integer.compare(num1, num2);
                    }
                }
                return Integer.compare(firstObject.length, secondObject.length)
            } catch (Exception e) {
                log.error("Ошибка сравнения орг. единиц: {}", e.getMessage())
                return 0
            }
        }
    }

    /**
     * Класс нужен для создания определений связей на модели БД Газпром на QAS, полученной через миграцию.
     * В этой модели вообще нет определений связи, а по ним было бы удобно и правильно обходить орг. структуру
     */
    class RelationCreator {
        void createRelationsInForest(Map<String, OrgUnitNode> rootNodes) {
            EdgesApi edgesApi = context.getApi(EdgesApi.class)
            TreeApi treeApi = context.getApi(TreeApi.class)
            for (OrgUnitNode rootNode : rootNodes.values()) {
//                deleteRelations(edgesApi, treeApi, rootNode)
                createRelations(edgesApi, rootNode)
            }
        }


        private void createRelations(EdgesApi edgesApi, OrgUnitNode node) {
            for (OrgUnitNode child : node.children) {
                createEdge(edgesApi, node, child)
                createRelations(edgesApi, child)
            }
        }

        private void deleteRelations(EdgesApi edgesApi, TreeApi treeApi, OrgUnitNode node) {
            deleteEdge(edgesApi, treeApi, node)
            for (OrgUnitNode child : node.children) {
                deleteRelations(edgesApi, treeApi, child)
            }
        }

        private void deleteEdge(EdgesApi edgesApi, TreeApi treeApi, OrgUnitNode source) {
            EdgeDefinitionSearch eds = new EdgeDefinitionSearch()
            eds.setRepositoryId(source.repositoryId)
            eds.setStartObjectsIds(Collections.singletonList(source.id))
            def searchResult = edgesApi.search(eds)

            Constructor<EraseRequest> erc = EraseRequest.class.getDeclaredConstructor(String.class, List.class)
            erc.setAccessible(true)

            searchResult.parallelStream().forEach {
                treeApi.deleteNode(it.nodeId.repositoryId, it.nodeId.id)
            }

            def eraseList = searchResult.stream()
                    .map { it.getNodeId().getId() }
                    .toList()

            EraseRequest er = erc.newInstance(source.repositoryId, eraseList)
            treeApi.erase(er)
        }


        private void createEdge(EdgesApi edgesApi, OrgUnitNode source, OrgUnitNode target) {
            EdgeDefinitionSearch eds = new EdgeDefinitionSearch()
            eds.setRepositoryId(source.repositoryId)
            eds.setStartObjectsIds(Collections.singletonList(source.id))
            eds.setEndObjectsIds(Collections.singletonList(target.id))
            def searchResult = edgesApi.search(eds)

            if (searchResult.isEmpty()) {
                EdgeDefinitionNode edge = new EdgeDefinitionNode()
                edge.setName("")
                edge.setParentNodeId(new NodeId(source.id, source.repositoryId, null))
                edge.setSourceObjectDefinitionId(source.id)
                edge.setTargetObjectDefinitionId(target.id)
                edge.setType(NodeType.EDGE)
                edge.setEdgeTypeId(target.isPosition ? "CT_IS_CRT_BY" : "CT_IS_SUPERIOR_1")
                EdgeDefinitionNode newEdge = edgesApi.create(edge)
                log.info("создана новая связь [{}] между объектами [{}->{}]", newEdge.nodeId.id, source.name, target.name)
            } else {
                def sr = searchResult.get(0)
                log.info("Найдена существующая связь [{}] между объектами [{}->{}]", sr.nodeId.id, source.name, target.name)
            }
        }
    }

    @Override
    void setContext(CustomScriptContext context) {
        this.context = context
    }
}