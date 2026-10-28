/**
 * ==============================================================================
 * FORMULARIO DE REGISTRO DE INSTALACIÓN DE CANDADOS INTELIGENTES
 * Luz del Sur - Lógica de Interfaz de Usuario (Frontend Puro)
 * ==============================================================================
 */

// ==============================================================================
// 1. CONFIGURACIÓN Y CATÁLOGO DE SUBESTACIONES (SED)
// ==============================================================================
/**
 * Catálogo de Subestaciones de Distribución (SED).
 * Esta constante contiene todas las opciones iniciales provistas.
 * Más adelante se puede reemplazar o alimentar fácilmente mediante una llamada
 * a la lista correspondiente de SharePoint (REST API o Microsoft Graph).
 */
const SED_LIST = [
  "00837S", "00676S", "00283S", "00379S", "01275S", "01975S", "00972S", "01981S",
  "01595S", "00804S", "01128S", "00767S", "01260S", "00670S", "00946S", "00816S",
  "01155S", "01156S", "00965S", "01703S", "00202S", "00071S", "00221S", "01629S",
  "00067S", "00222S", "00081S", "01001S", "01215S", "00596S", "00630S", "01909S",
  "00789S", "00623S", "00383S", "00903S", "01646S", "00266S", "01408S", "00407S",
  "00356S", "00702S", "01006S", "01022S", "00599S", "01288S", "00768S", "01952S",
  "00811S", "01835S", "01480S", "01239S", "01427S", "01959S", "00457S", "00586S",
  "00030S", "00814S", "00278S", "00132S", "00162S", "01109S", "01436S", "00352S",
  "00169S", "00021S", "01033S", "00157S", "00031S", "00147S", "01550S", "01419S",
  "01855S", "00863S", "01797S", "00889S", "01856S", "01370S", "01162S", "01683S",
  "00458S", "01171S", "01960S", "00671S", "01441S", "01268S", "01755S", "00945S",
  "01295S", "00940S", "00536S", "01917S", "01059S", "01995S", "01946S", "01306S",
  "00981S", "01847S", "01414S", "01349S", "01587S", "00839S", "01411S", "00849S",
  "00845S", "01859S", "01438S", "01843S", "01838S", "01822S", "00800S", "01546S",
  "00910S", "01385S", "00133S", "00320S", "00392S", "00113S", "01861S", "00310S",
  "00482S", "00349S", "01205S", "01969S", "00337S", "00904S", "01373S", "01312S",
  "00317S", "01898S", "00146S", "01357S", "01346S", "01333S", "01430S", "01966S",
  "01300S", "00971S", "00759S", "00727S", "00490S", "00303S", "00744S", "00351S",
  "00751S", "00728S", "01289S", "01726S", "01361S", "01653S", "00256S", "01551S",
  "01978S", "01725S", "01169S", "00209S", "00348S", "00072S", "01170S", "00251S",
  "01964S", "00554S", "00425S", "00284S", "00621S", "00518S", "00376S", "00220S",
  "00380S", "00193S", "00883S", "00247S", "00795S", "00063S", "00405S", "00342S",
  "00064S", "00487S", "00985S", "00643S", "00602S", "01282S", "00968S", "01645S",
  "00601S", "00741S", "01281S", "00499S", "01041S", "00790S", "01047S", "01121S",
  "00696S", "00473S", "00672S", "01366S", "00994S", "00561S", "00537S", "01313S",
  "01098S", "01292S", "01574S", "01194S", "01086S", "01016S", "00213S", "00439S",
  "01452S", "01460S", "01463S", "01470S", "01476S", "01504S", "01600S", "01359S",
  "01515S", "01522S", "01523S", "01566S", "01572S", "01580S", "01581S", "01602S",
  "01611S", "01643S", "01665S", "1384BS", "01668S", "01684S", "01719S", "80001S",
  "01742S", "01746S", "01778S", "01757S", "00229S", "01718S", "01794S", "01807S",
  "01820S", "01815S", "01850S", "01853S", "01879S", "01882S", "01890S", "01907S",
  "01933S", "01938S", "01962S", "01982S", "01973S", "01988S", "01483S", "1394DS",
  "1383BS", "01258S", "00852S", "00933S", "00578S", "01618S", "01274S", "01593S",
  "00475S", "01744S", "01864S", "01537S", "01943S", "01273S", "01863S", "01656S",
  "00474S", "00974S", "01321S", "01956S", "01710S", "01101S", "00466S", "00187S",
  "00562S", "01222S", "01112S", "00312S", "00070S", "00915S", "00698S", "00842S",
  "01401S", "01369S", "00381S", "01832S", "01305S", "01771S", "01406S", "01066S",
  "00459S", "00406S", "00269S", "00270S", "01164S", "01007S", "01141S", "01745S",
  "01950S", "00705S", "01261S", "00632S", "01469S", "01365S", "01368S", "00812S",
  "01032S", "00131S", "00329S", "00137S", "01310S", "01925S", "00120S", "01043S",
  "00429S", "00049S", "00117S", "00127S", "00089S", "00300S", "01031S", "00978S",
  "01663S", "01831S", "01559S", "01830S", "01379S", "01354S", "01997S", "01913S",
  "01784S", "01431S", "00491S", "00714S", "00563S", "01091S", "01711S", "01394S",
  "01635S", "01445S", "01203S", "00941S", "01304S", "01317S", "01052S", "00663S",
  "01145S", "00404S", "00558S", "01624S", "01478S", "01413S", "01389S", "01403S",
  "01770S", "00848S", "00900S", "01210S", "01768S", "01439S", "01844S", "01839S",
  "01821S", "01294S", "00928S", "01560S", "01416S", "00888S", "00051S", "01386S",
  "00153S", "00884S", "01481S", "00336S", "00869S", "00906S", "00167S", "01374S",
  "00206S", "01371S", "01376S", "00056S", "00418S", "01899S", "01672S", "00531S",
  "01945S", "00506S", "01616S", "00982S", "01301S", "01097S", "01053S", "01185S",
  "00305S", "00302S", "00308S", "00894S", "00752S", "00600S", "01631S", "01637S",
  "01038S", "01420S", "01998S", "01761S", "00866S", "01894S", "01188S", "00834S",
  "01077S", "00136S", "01722S", "00065S", "00412S", "00555S", "00334S", "00232S",
  "01028S", "00078S", "00367S", "00912S", "01064S", "00077S", "01409S", "00184S",
  "01789S", "01021S", "00074S", "00073S", "00830S", "00469S", "01116S", "00644S",
  "00642S", "01626S", "01810S", "01197S", "01586S", "00740S", "01796S", "00538S",
  "01298S", "00687S", "00697S", "00471S", "01248S", "00995S", "00583S", "00508S",
  "01666S", "01071S", "00879S", "01315S", "01228S", "01613S", "00796S", "01195S",
  "01085S", "00762S", "00622S", "00797S", "01454S", "01461S", "01466S", "01472S",
  "01487S", "01525S", "01540S", "1394BS", "01520S", "01541S", "01555S", "01567S",
  "01577S", "01582S", "01598S", "01609S", "01620S", "01630S", "01657S", "01678S",
  "01677S", "01709S", "01705S", "01673S", "01723S", "01754S", "01769S", "01766S",
  "05050S", "01799S", "01792S", "01801S", "01819S", "01827S", "01804S", "01867S",
  "01880S", "01885S", "01858S", "01908S", "01935S", "01948S", "01949S", "01965S",
  "01970S", "01989S", "01318S", "01291S", "01265S", "01262S", "00542S", "00411S",
  "01874S", "01332S", "01422S", "01679S", "00774S", "01249S", "01325S", "01027S",
  "01026S", "01872S", "01897S", "01442S", "01279S", "00189S", "00911S", "00448S",
  "00145S", "01752S", "00208S", "00314S", "00414S", "00172S", "00360S", "00523S",
  "00914S", "00917S", "01017S", "00617S", "01425S", "00618S", "00273S", "00477S",
  "01144S", "00265S", "01532S", "01531S", "00331S", "00582S", "00708S", "00382S",
  "01163S", "01093S", "00690S", "00771S", "01736S", "01686S", "00706S", "00525S",
  "01834S", "01129S", "01570S", "01039S", "00023S", "00979S", "00330S", "00155S",
  "00453S", "00291S", "00036S", "01429S", "00454S", "00044S", "00054S", "00168S",
  "01696S", "00268S", "01034S", "00224S", "01660S", "01829S", "01008S", "01634S",
  "01381S", "01347S", "01428S", "01449S", "01805S", "00922S", "00480S", "01606S",
  "01356S", "01994S", "01793S", "01174S", "01603S", "01798S", "01202S", "01448S",
  "00809S", "00588S", "01050S", "00664S", "01846S", "00456S", "00808S", "01968S",
  "01402S", "01547S", "01390S", "01404S", "01405S", "01494S", "01440S", "01209S",
  "01211S", "01842S", "01845S", "01840S", "01104S", "01230S", "01057S", "01765S",
  "01446S", "01485S", "01397S", "01556S", "00058S", "00462S", "00333S", "00140S",
  "01311S", "00415S", "01873S", "00200S", "00397S", "01410S", "01387S", "00288S",
  "00175S", "00059S", "01087S", "00530S", "01327S", "00645S", "00597S", "01930S",
  "01299S", "01364S", "00936S", "00460S", "00304S", "00306S", "00589S", "01302S",
  "00689S", "00846S", "01632S", "01790S", "01601S", "01869S", "02000S", "01172S",
  "01375S", "01160S", "01893S", "00201S", "01947S", "00181S", "00987S", "01748S",
  "00061S", "00476S", "00516S", "00479S", "00442S", "00246S", "00218S", "01812S",
  "00075S", "00151S", "01353S", "01107S", "00066S", "00451S", "01377S", "00766S",
  "01142S", "01817S", "00649S", "00603S", "00733S", "00736S", "01811S", "00648S",
  "01213S", "00489S", "01535S", "01182S", "01123S", "00575S", "00581S", "00210S",
  "00564S", "01099S", "00550S", "00551S", "00916S", "00673S", "00886S", "00734S",
  "01308S", "01276S", "00591S", "01243S", "00788S", "00763S", "00470S", "01447S",
  "01457S", "01462S", "01464S", "01473S", "01496S", "01497S", "01539S", "01501S",
  "01517S", "01521S", "01562S", "01576S", "01578S", "01585S", "01599S", "01610S",
  "01641S", "01651S", "01676S", "01697S", "01713S", "01720S", "01738S", "01747S",
  "01753S", "01760S", "01781S", "06000S", "01788S", "01802S", "04030S", "01825S",
  "01828S", "01823S", "01854S", "01878S", "01886S", "01895S", "01905S", "01934S",
  "01954S", "01976S", "01967S", "01983S", "01498S", "01316S", "01267S", "01277S",
  "01775S", "00844S", "00409S", "00878S", "01286S", "00973S", "01348S", "00785S",
  "00999S", "00743S", "01638S", "01259S", "00913S", "01592S", "00552S", "00847S",
  "00826S", "00511S", "01451S", "00148S", "01491S", "00503S", "00540S", "00079S",
  "00236S", "00068S", "00934S", "00935S", "01166S", "01246S", "01037S", "00619S",
  "01527S", "00556S", "01809S", "00362S", "00335S", "00547S", "00297S", "01067S",
  "00896S", "01984S", "01216S", "00560S", "00688S", "00769S", "01951S", "00041S",
  "00707S", "01465S", "00753S", "00909S", "01283S", "00029S", "00436S", "00433S",
  "00323S", "00661S", "00052S", "00161S", "00823S", "00301S", "00322S", "00170S",
  "01103S", "01435S", "00043S", "00028S", "00194S", "01380S", "01049S", "01944S",
  "01503S", "01418S", "01857S", "00820S", "00827S", "01549S", "00862S", "00481S",
  "01264S", "01137S", "01395S", "01824S", "01303S", "01619S", "01012S", "01297S",
  "01655S", "00535S", "00510S", "00667S", "00665S", "01151S", "01045S", "00807S",
  "01266S", "01911S", "01548S", "01391S", "00843S", "01412S", "01774S", "01095S",
  "01212S", "01437S", "01533S", "01837S", "01841S", "00786S", "01400S", "01787S",
  "01767S", "00854S", "00850S", "00115S", "00324S", "00154S", "01860S", "00271S",
  "01749S", "01293S", "00465S", "00289S", "00824S", "00204S", "01372S", "01399S",
  "00055S", "00296S", "00188S", "01417S", "00426S", "01040S", "01808S", "01186S",
  "01189S", "00633S", "00760S", "00773S", "00624S", "01350S", "00307S", "01102S",
  "01528S", "00729S", "01543S", "01453S", "01360S", "01654S", "01800S", "01999S",
  "01700S", "00840S", "01168S", "00141S", "01493S", "01426S", "00361S", "00186S",
  "00895S", "00964S", "00389S", "00452S", "00313S", "00277S", "00434S", "01490S",
  "01910S", "00183S", "00967S", "00142S", "00226S", "00062S", "01806S", "01120S",
  "01004S", "00488S", "00984S", "00647S", "00650S", "00828S", "01339S", "00653S",
  "00646S", "01214S", "01977S", "01649S", "00557S", "00791S", "00988S", "00387S",
  "00472S", "01724S", "00421S", "00549S", "01396S", "00559S", "01024S", "01314S",
  "00801S", "01309S", "01706S", "01920S", "01530S", "00787S", "00620S", "00438S",
  "01444S", "01458S", "01455S", "01468S", "01474S", "01433S", "01508S", "01423S",
  "01506S", "01524S", "01554S", "01564S", "01575S", "01579S", "01588S", "01604S",
  "01607S", "01628S", "01636S", "1388BS", "01669S", "01693S", "01695S", "01727S",
  "01743S", "01750S", "01756S", "01739S", "01051S", "04010S", "01795S", "01791S",
  "01816S", "01818S", "01848S", "01852S", "01876S", "01877S", "01889S", "01803S",
  "01924S", "01939S", "01915S", "01974S", "01972S", "01902S", "01992S", "01307S",
  "01263S", "01271S", "01927S", "01928S", "01931S", "01384S", "01383S", "1417BS",
  "01388S", "01214S", "01958S", "01231S", "CEPS90000", "CEPS90002", "CEPS90007",
  "CEPS90014", "CEPS90036", "CEPS90038", "CEPS90043", "CEPS90047", "CEPS3010",
  "01225S", "01229S"
];

// ==============================================================================
// 2. REGLAS DE VALIDACIÓN Y PARÁMETROS CONFIGURABLES
// ==============================================================================
/**
 * Expresión regular para validar el código de candado a instalar.
 * Configuración actual: Letras y números sin espacios, mínimo 8 caracteres.
 * Modificar aquí si las reglas de nomenclatura de candados cambian en el futuro.
 */
const REGEX_CODIGO_CANDADO = /^[A-Z0-9]{8,}$/;

/**
 * Límite máximo de fotos o evidencias permitidas.
 */
const MAX_FOTOS = 10;

/**
 * Endpoint del Webhook de Power Automate para registro en la lista de SharePoint.
 */
const URL_POWER_AUTOMATE = "https://default1c0051dd45964b1a9849d060735057.69.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/23/workflows/5effa9baaef34ba7aeb4f8e1c1cbda1a/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=JOk4GlmUXy47A0-zN5UIcVrRKNXQJLcK6BhGWFzBESo";

// ==============================================================================
// 3. ESTADO GLOBAL DE LA APLICACIÓN
// ==============================================================================
const state = {
  selectedSed: "",
  fotos: [], // Array de objetos { id, file, url, name, size }
  highlightedIndex: -1
};

// ==============================================================================
// 4. ELEMENTOS DEL DOM
// ==============================================================================
const form = document.getElementById("formRegistro");
const btnReset = document.getElementById("btnReset");

// SED Elements (Unificado)
const sedInput = document.getElementById("sedInput");
const sedOptionsList = document.getElementById("sedOptionsList");
const sedClearBtn = document.getElementById("sedClearBtn");
const sedHiddenInput = document.getElementById("sed");

// Lock type, Code, Date & Responsible
const tipoCheckboxes = document.querySelectorAll('input[name="tipoCandado"]');
const codigoInput = document.getElementById("codigoCandado");
const observacionesInput = document.getElementById("observaciones");
const responsableInput = document.getElementById("responsable");
const fechaInput = document.getElementById("fechaInstalacion");

// Photo Elements
const fotoInput = document.getElementById("fotoInput");
const btnUploadTrigger = document.getElementById("btnUploadTrigger");
const uploadDropzone = document.getElementById("uploadDropzone");
const photoPreviewGrid = document.getElementById("photoPreviewGrid");
const photoCounterBadge = document.getElementById("photoCounterBadge");
const photoStatusText = document.getElementById("photoStatusText");

// Modal Elements
const modalExito = document.getElementById("modalExito");
const btnModalNuevo = document.getElementById("btnModalNuevo");
const btnModalCerrar = document.getElementById("btnModalCerrar");
const summarySed = document.getElementById("summarySed");
const summaryTipo = document.getElementById("summaryTipo");
const summaryCodigo = document.getElementById("summaryCodigo");
const summaryResponsable = document.getElementById("summaryResponsable");
const summaryFecha = document.getElementById("summaryFecha");
const summaryObservaciones = document.getElementById("summaryObservaciones");
const summaryFotosCount = document.getElementById("summaryFotosCount");
const summaryThumbsGrid = document.getElementById("summaryThumbsGrid");

// ==============================================================================
// 5. INICIALIZACIÓN
// ==============================================================================
document.addEventListener("DOMContentLoaded", () => {
  initDateInput();
  initSedDropdown();
  initLockTypes();
  initCodeInput();
  initPhotoUploader();
  initFormValidation();
  initModal();
});

/**
 * Inicializa el campo de fecha con la fecha local de hoy (YYYY-MM-DD).
 */
function initDateInput() {
  if (fechaInput && !fechaInput.value) {
    const hoy = new Date();
    const year = hoy.getFullYear();
    const month = String(hoy.getMonth() + 1).padStart(2, "0");
    const day = String(hoy.getDate()).padStart(2, "0");
    fechaInput.value = `${year}-${month}-${day}`;
  }
}

// ==============================================================================
// ==============================================================================
// 6. COMPONENTE: SELECTOR DE SED (UNIFICADO)
// ==============================================================================
function initSedDropdown() {
  // Al hacer clic o foco en el input, muestra las opciones
  sedInput.addEventListener("focus", () => {
    filtrarYMostrarSed(sedInput.value.trim());
  });

  // Al escribir, filtra en tiempo real
  sedInput.addEventListener("input", () => {
    sedInput.value = sedInput.value.toUpperCase();
    const query = sedInput.value.trim();
    sedClearBtn.style.display = query ? "block" : "none";
    filtrarYMostrarSed(query);
  });

  // Botón para limpiar
  sedClearBtn.addEventListener("click", () => {
    clearSedSelection();
    sedInput.focus();
  });

  // Navegación por teclado (Flechas y Enter)
  sedInput.addEventListener("keydown", handleSedKeydown);

  // Cerrar si hace clic fuera
  document.addEventListener("click", (e) => {
    const container = document.getElementById("sedSelectContainer");
    if (!container.contains(e.target)) {
      sedOptionsList.classList.remove("open");
      validarSedEscrita();
    }
  });
}

function filtrarYMostrarSed(query) {
  const filtradas = query 
    ? SED_LIST.filter(item => item.toUpperCase().includes(query.toUpperCase()))
    : SED_LIST;

  renderSedOptions(filtradas, query);
  sedOptionsList.classList.add("open");
}

function renderSedOptions(list, query = "") {
  sedOptionsList.innerHTML = "";
  state.highlightedIndex = -1;

  if (list.length === 0) {
    const li = document.createElement("li");
    li.className = "dropdown-no-results";
    li.textContent = `No se encontró "${query}"`;
    sedOptionsList.appendChild(li);
    return;
  }

  list.forEach(item => {
    const li = document.createElement("li");
    li.className = "dropdown-option";
    li.dataset.value = item;
    if (item === state.selectedSed) li.classList.add("selected");

    if (query) {
      const regex = new RegExp(`(${query})`, "gi");
      li.innerHTML = `<span class="option-text">${item.replace(regex, "<strong style='color:#F37021;'>$1</strong>")}</span>`;
    } else {
      li.innerHTML = `<span class="option-text">${item}</span>`;
    }

    li.addEventListener("click", () => selectSed(item));
    sedOptionsList.appendChild(li);
  });
}

function selectSed(sedValue) {
  state.selectedSed = sedValue;
  sedHiddenInput.value = sedValue;
  sedInput.value = sedValue;
  sedClearBtn.style.display = "block";
  sedOptionsList.classList.remove("open");
  clearFieldError("sed");
}

function clearSedSelection() {
  state.selectedSed = "";
  sedHiddenInput.value = "";
  sedInput.value = "";
  sedClearBtn.style.display = "none";
  sedOptionsList.classList.remove("open");
}

function validarSedEscrita() {
  const val = sedInput.value.trim().toUpperCase();
  if (!val) {
    clearSedSelection();
  } else if (SED_LIST.includes(val)) {
    selectSed(val);
  } else {
    // Si escribió algo que no existe en la lista oficial
    state.selectedSed = "";
    sedHiddenInput.value = "";
  }
}

function handleSedKeydown(e) {
  const options = sedOptionsList.querySelectorAll(".dropdown-option");
  if (!sedOptionsList.classList.contains("open") || options.length === 0) return;

  if (e.key === "ArrowDown") {
    e.preventDefault();
    state.highlightedIndex = (state.highlightedIndex + 1) % options.length;
    updateHighlightedOption(options);
  } else if (e.key === "ArrowUp") {
    e.preventDefault();
    state.highlightedIndex = (state.highlightedIndex - 1 + options.length) % options.length;
    updateHighlightedOption(options);
  } else if (e.key === "Enter") {
    e.preventDefault();
    if (state.highlightedIndex >= 0 && options[state.highlightedIndex]) {
      selectSed(options[state.highlightedIndex].dataset.value);
    } else if (options.length === 1) {
      selectSed(options[0].dataset.value);
    }
  } else if (e.key === "Escape") {
    sedOptionsList.classList.remove("open");
  }
}

function updateHighlightedOption(options) {
  options.forEach((opt, idx) => {
    if (idx === state.highlightedIndex) {
      opt.classList.add("highlighted");
      opt.scrollIntoView({ block: "nearest" });
    } else {
      opt.classList.remove("highlighted");
    }
  });
}

// ==============================================================================
// 7. COMPONENTE: TIPO DE CANDADO (Selección Múltiple)
// ==============================================================================
function initLockTypes() {
  tipoCheckboxes.forEach((checkbox) => {
    const label = checkbox.closest(".lock-chip-label");

    // Sincronizar estado visual
    const syncCheckedState = () => {
      if (checkbox.checked) {
        label.classList.add("is-checked");
      } else {
        label.classList.remove("is-checked");
      }
      clearFieldError("tipoCandado");
    };

    checkbox.addEventListener("change", syncCheckedState);
  });
}

// ==============================================================================
// 8. COMPONENTE: CÓDIGO DE CANDADO (Mayúsculas y Regex)
// ==============================================================================
function initCodeInput() {
  // Convertir a mayúsculas automáticamente en tiempo real
  codigoInput.addEventListener("input", () => {
    const cursorPosition = codigoInput.selectionStart;
    codigoInput.value = codigoInput.value.toUpperCase().replace(/\s+/g, "");
    codigoInput.setSelectionRange(cursorPosition, cursorPosition);

    if (codigoInput.value.length > 0) {
      validateCodigoCandado(false);
    } else {
      clearFieldError("codigoCandado");
    }
  });

  codigoInput.addEventListener("blur", () => {
    if (codigoInput.value.trim().length > 0) {
      validateCodigoCandado(true);
    }
  });
}

function validateCodigoCandado(showAlert = true) {
  const val = codigoInput.value.trim();
  if (!val) {
    if (showAlert) setFieldError("codigoCandado", "El código de candado es obligatorio.");
    return false;
  }

  if (!REGEX_CODIGO_CANDADO.test(val)) {
    if (showAlert) {
      setFieldError(
        "codigoCandado",
        "Formato inválido. Debe contener mínimo 9 caracteres alfanuméricos sin espacios (ej. Ejem:L0000xxxx)."
      );
    }
    return false;
  }

  clearFieldError("codigoCandado");
  return true;
}

// ==============================================================================
// 9. COMPONENTE: FOTOS O EVIDENCIA (Múltiple, Cámara/Galería, Drag & Drop)
// ==============================================================================
function initPhotoUploader() {
  // Disparar input de archivo al hacer click en el botón o dropzone
  btnUploadTrigger.addEventListener("click", (e) => {
    e.stopPropagation();
    fotoInput.click();
  });

  uploadDropzone.addEventListener("click", () => {
    fotoInput.click();
  });

  // Accesibilidad por teclado en la dropzone
  uploadDropzone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      fotoInput.click();
    }
  });

  // Manejo de selección mediante selector de archivos
  fotoInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFiles(Array.from(e.target.files));
      fotoInput.value = ""; // Permite seleccionar el mismo archivo de nuevo si se borró
    }
  });

  // Drag and drop para equipos de escritorio
  ["dragenter", "dragover"].forEach(eventName => {
    uploadDropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      uploadDropzone.classList.add("drag-over");
    });
  });

  ["dragleave", "drop"].forEach(eventName => {
    uploadDropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      uploadDropzone.classList.remove("drag-over");
    });
  });

  uploadDropzone.addEventListener("drop", (e) => {
    if (e.dataTransfer && e.dataTransfer.files) {
      handleFiles(Array.from(e.dataTransfer.files));
    }
  });
}

function handleFiles(files) {
  clearFieldError("fotos");

  // Filtrar solo imágenes
  const imageFiles = files.filter(file => file.type.startsWith("image/"));

  if (imageFiles.length === 0) {
    setFieldError("fotos", "Solo se admiten archivos de imagen (JPG, PNG, WEBP, etc.).");
    return;
  }

  const espacioDisponible = MAX_FOTOS - state.fotos.length;

  if (espacioDisponible <= 0) {
    setFieldError("fotos", `Ya alcanzaste el límite máximo de ${MAX_FOTOS} fotos.`);
    return;
  }

  let fotosAñadir = imageFiles;
  if (imageFiles.length > espacioDisponible) {
    fotosAñadir = imageFiles.slice(0, espacioDisponible);
    setFieldError(
      "fotos",
      `Se agregaron solo ${espacioDisponible} foto(s). Límite máximo de ${MAX_FOTOS} fotos alcanzado.`
    );
  }

  // Procesar fotos y generar URLs seguras para miniaturas
  fotosAñadir.forEach(file => {
    const fotoItem = {
      id: "foto_" + Date.now() + "_" + Math.random().toString(36).substr(2, 6),
      file: file,
      url: URL.createObjectURL(file),
      name: file.name,
      size: formatFileSize(file.size)
    };
    state.fotos.push(fotoItem);
  });

  renderPhotoPreviews();
}

function removePhoto(id) {
  const index = state.fotos.findIndex(f => f.id === id);
  if (index !== -1) {
    // Liberar memoria del ObjectURL
    URL.revokeObjectURL(state.fotos[index].url);
    state.fotos.splice(index, 1);
    renderPhotoPreviews();

    if (state.fotos.length > 0) {
      clearFieldError("fotos");
    }
  }
}

function renderPhotoPreviews() {
  photoPreviewGrid.innerHTML = "";
  const count = state.fotos.length;

  // Actualizar contador y estado
  photoCounterBadge.textContent = `${count} / ${MAX_FOTOS} fotos`;

  if (count >= MAX_FOTOS) {
    photoCounterBadge.classList.add("limit-reached");
    photoStatusText.textContent = "Límite máximo alcanzado";
  } else {
    photoCounterBadge.classList.remove("limit-reached");
    photoStatusText.textContent = count === 0 ? "Mínimo 1 foto requerida" : `${MAX_FOTOS - count} disponible(s)`;
  }

  state.fotos.forEach(foto => {
    const card = document.createElement("div");
    card.className = "photo-preview-card";
    card.dataset.id = foto.id;

    card.innerHTML = `
      <div class="photo-thumb-container">
        <img src="${foto.url}" alt="${escapeHtml(foto.name)}" class="photo-thumb-img" loading="lazy">
        <button type="button" class="btn-remove-photo" title="Eliminar foto" aria-label="Eliminar foto ${escapeHtml(foto.name)}">
          &times;
        </button>
      </div>
      <div class="photo-info-bar">
        <span class="photo-name" title="${escapeHtml(foto.name)}">${escapeHtml(foto.name)}</span>
        <span class="photo-size">${foto.size}</span>
      </div>
    `;

    card.querySelector(".btn-remove-photo").addEventListener("click", () => {
      removePhoto(foto.id);
    });

    photoPreviewGrid.appendChild(card);
  });
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

// ==============================================================================
// 10. VALIDACIÓN Y GESTIÓN DE ERRORES EN TIEMPO REAL
// ==============================================================================
function initFormValidation() {
  // Validación en tiempo real al escribir en campos de texto
  responsableInput.addEventListener("input", () => {
    if (responsableInput.value.trim().length > 0) {
      clearFieldError("responsable");
    }
  });

  responsableInput.addEventListener("blur", () => {
    if (!responsableInput.value.trim()) {
      setFieldError("responsable", "Debes ingresar el nombre del responsable.");
    }
  });

  // Envío del formulario
  form.addEventListener("submit", async (e) => {
    e.preventDefault();

    const esValido = validateAllFields();

    if (!esValido) {
      scrollToFirstError();
      return;
    }

    // Preparar objeto ordenado con todos los datos
    const selectedTipos = Array.from(tipoCheckboxes)
      .filter(cb => cb.checked)
      .map(cb => cb.value);

    const payload = {
      timestamp: new Date().toISOString(),
      fechaFormateada: new Date().toLocaleString("es-PE"),
      sed: state.selectedSed,
      tipoCandado: selectedTipos,
      codigoCandado: codigoInput.value.trim().toUpperCase(),
      responsable: responsableInput.value.trim(),
      fechaInstalacion: fechaInput ? fechaInput.value : new Date().toISOString().split("T")[0],
      observaciones: observacionesInput.value.trim() || "Sin observaciones",
      cantidadFotos: state.fotos.length,
      fotos: state.fotos.map((f, index) => ({
        indice: index + 1,
        nombreArchivo: f.name,
        tamanoLegible: f.size,
        tipoMime: f.file.type,
        // Referencia al objeto File real para la futura carga en SharePoint
        fileRef: f.file
      }))
    };

    // Estado de carga en el botón durante el envío a Power Automate / SharePoint
    const btnSubmit = document.getElementById("btnSubmit");
    const textoOriginalBtn = btnSubmit.innerHTML;
    btnSubmit.disabled = true;
    btnSubmit.innerHTML = '<span>⏳</span> Enviando a SharePoint...';

    try {
      // Envío real a Power Automate con Base64
      await enviarRegistro(payload);

      // Mostrar modal de éxito con el resumen
      showSuccessModal(payload);
    } catch (error) {
      console.error("Error al transmitir a SharePoint / Power Automate:", error);
      alert("Hubo un inconveniente al enviar a SharePoint:\n" + error.message + "\n\nVerifica que tu flujo de Power Automate esté activado.");
    } finally {
      btnSubmit.disabled = false;
      btnSubmit.innerHTML = textoOriginalBtn;
    }
  });

  // Botón de limpiar formulario
  btnReset.addEventListener("click", () => {
    const tieneDatos = state.selectedSed || 
                       codigoInput.value.trim() || 
                       responsableInput.value.trim() || 
                       observacionesInput.value.trim() || 
                       state.fotos.length > 0;

    if (tieneDatos) {
      const confirmacion = confirm("¿Deseas limpiar todos los campos del formulario?");
      if (confirmacion) {
        resetCompleteForm();
      }
    } else {
      resetCompleteForm();
    }
  });
}

function validateAllFields() {
  let isValid = true;

  // 1. Validar SED
  if (!state.selectedSed) {
    setFieldError("sed", "Debes seleccionar una subestación (SED) de la lista.");
    isValid = false;
  } else {
    clearFieldError("sed");
  }

  // 2. Validar Tipo de Candado
  const algunTipoSeleccionado = Array.from(tipoCheckboxes).some(cb => cb.checked);
  if (!algunTipoSeleccionado) {
    setFieldError("tipoCandado", "Debes seleccionar al menos un tipo de candado (MULTILOCK o YARDENI).");
    isValid = false;
  } else {
    clearFieldError("tipoCandado");
  }

  // 3. Validar Código de Candado
  if (!validateCodigoCandado(true)) {
    isValid = false;
  }

  // 4. Validar Responsable
  if (!responsableInput.value.trim()) {
    setFieldError("responsable", "Debes ingresar el nombre del responsable de la instalación.");
    isValid = false;
  } else {
    clearFieldError("responsable");
  }

  // 5. Validar Fecha de Instalación
  if (fechaInput && !fechaInput.value) {
    setFieldError("fecha", "Debes seleccionar la fecha de instalación.");
    isValid = false;
  } else {
    clearFieldError("fecha");
  }

  // 6. Validar Fotos
  if (state.fotos.length === 0) {
    setFieldError("fotos", "Debes adjuntar al menos una foto como evidencia de la instalación.");
    isValid = false;
  } else if (state.fotos.length > MAX_FOTOS) {
    setFieldError("fotos", `Se excede el máximo permitido de ${MAX_FOTOS} fotos.`);
    isValid = false;
  } else {
    clearFieldError("fotos");
  }

  return isValid;
}

function setFieldError(fieldId, message) {
  const group = document.getElementById(`group-${fieldId}`);
  const errorEl = document.getElementById(`error-${fieldId}`);
  if (group) group.classList.add("has-error");
  if (errorEl) {
    errorEl.textContent = message;
    errorEl.style.display = "flex";
  }
}

function clearFieldError(fieldId) {
  const group = document.getElementById(`group-${fieldId}`);
  const errorEl = document.getElementById(`error-${fieldId}`);
  if (group) group.classList.remove("has-error");
  if (errorEl) {
    errorEl.textContent = "";
    errorEl.style.display = "none";
  }
}

function scrollToFirstError() {
  const firstError = document.querySelector(".form-group.has-error");
  if (firstError) {
    firstError.scrollIntoView({ behavior: "smooth", block: "center" });
    const focusable = firstError.querySelector("input:not([type=hidden]), textarea, button, [tabindex='0']");
    if (focusable) {
      focusable.focus();
    }
  }
}

function resetCompleteForm() {
  form.reset();

  // Restaurar fecha de hoy
  initDateInput();

  // Limpiar SED
  clearSedSelection();

  // Desmarcar chips de candado
  document.querySelectorAll(".lock-chip-label").forEach(l => l.classList.remove("is-checked"));

  // Limpiar fotos
  state.fotos.forEach(f => URL.revokeObjectURL(f.url));
  state.fotos = [];
  renderPhotoPreviews();

  // Limpiar errores visuales
  ["sed", "tipoCandado", "codigoCandado", "responsable", "fecha", "fotos"].forEach(fieldId => {
    clearFieldError(fieldId);
  });

  window.scrollTo({ top: 0, behavior: "smooth" });
}

// ==============================================================================
// 11. FUNCIÓN DE ENVÍO DIRECTA A POWER AUTOMATE / SHAREPOINT
// ==============================================================================
/**
 * Convierte un archivo de imagen (File) a cadena Base64 pura
 * lista para decodificarse en Power Automate con base64ToBinary(...)
 */
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result;
      const base64Pure = dataUrl.split(",")[1];
      resolve(base64Pure);
    };
    reader.onerror = error => reject(error);
    reader.readAsDataURL(file);
  });
}

/**
 * Envío del registro de instalación a Power Automate.
 * Crea el elemento en la Lista de SharePoint y adjunta las fotos en Base64.
 */
async function enviarRegistro(datos) {
  console.group("🚀 [LUZ DEL SUR] TRANSMISIÓN A SHAREPOINT VÍA POWER AUTOMATE");
  console.log("Endpoint:", URL_POWER_AUTOMATE);

  // 1. Convertir todas las fotos adjuntas a Base64
  const fotosBase64 = await Promise.all(
    state.fotos.map(async (foto) => {
      const contentBytes = await fileToBase64(foto.file);
      return {
        id: foto.id,
        name: foto.name,
        status: "ready",
        sizeOriginalKB: Math.round(foto.file.size / 1024),
        sizeCompressedKB: Math.round(foto.file.size / 1024),
        contentBytes: contentBytes
      };
    })
  );

  // 2. Estructurar el paquete con los nombres que espera SharePoint
  const payloadSharePoint = {
    sed: datos.sed,
    tipoCandado: Array.isArray(datos.tipoCandado) ? datos.tipoCandado.join(", ") : datos.tipoCandado,
    codigoCandado: datos.codigoCandado,
    responsable: datos.responsable,
    fechaInstalacion: datos.fechaInstalacion + "T12:00:00",
    observaciones: datos.observaciones || "Sin observaciones",
    fechaRegistro: new Date().toLocaleString("sv-SE", { timeZone: "America/Lima" }).replace(" ", "T"),
    fotos: fotosBase64
  };

  console.log("Payload preparado (resumen):", {
    sed: payloadSharePoint.sed,
    tipoCandado: payloadSharePoint.tipoCandado,
    codigoCandado: payloadSharePoint.codigoCandado,
    responsable: payloadSharePoint.responsable,
    fechaInstalacion: payloadSharePoint.fechaInstalacion,
    observaciones: payloadSharePoint.observaciones,
    totalFotos: payloadSharePoint.fotos.length
  });

  // 3. Transmisión POST a Power Automate
  const response = await fetch(URL_POWER_AUTOMATE, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payloadSharePoint)
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => "");
    console.error("Respuesta fallida de Power Automate:", response.status, errorText);
    throw new Error(`Código ${response.status} en Power Automate: ${errorText || response.statusText}`);
  }

  console.log("✅ Transmisión exitosa. Código de respuesta:", response.status);
  console.groupEnd();

  let dataRespuesta = {};
  try {
    dataRespuesta = await response.json();
  } catch (e) {
    // Si la acción 'Respuesta' responde texto vacío o 200 sin JSON
  }
  return dataRespuesta;
}

// ==============================================================================
// 12. MODAL DE RESUMEN Y ÉXITO
// ==============================================================================
function initModal() {
  btnModalCerrar.addEventListener("click", closeModal);
  btnModalNuevo.addEventListener("click", () => {
    closeModal();
    resetCompleteForm();
  });

  // Cerrar al hacer click en el backdrop
  modalExito.addEventListener("click", (e) => {
    if (e.target === modalExito) {
      closeModal();
    }
  });

  // Cerrar con Escape
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && modalExito.classList.contains("active")) {
      closeModal();
    }
  });
}

function showSuccessModal(payload) {
  summarySed.textContent = payload.sed;
  summaryTipo.textContent = payload.tipoCandado.join(", ");
  summaryCodigo.textContent = payload.codigoCandado;
  summaryResponsable.textContent = payload.responsable;
  if (summaryFecha) summaryFecha.textContent = payload.fechaInstalacion;
  summaryObservaciones.textContent = payload.observaciones;
  summaryFotosCount.textContent = `${payload.cantidadFotos} foto(s) adjunta(s)`;

  // Renderizar miniaturas en el resumen
  summaryThumbsGrid.innerHTML = "";
  state.fotos.forEach(foto => {
    const thumb = document.createElement("div");
    thumb.className = "summary-thumb";
    thumb.innerHTML = `<img src="${foto.url}" alt="${escapeHtml(foto.name)}">`;
    summaryThumbsGrid.appendChild(thumb);
  });

  modalExito.classList.add("active");
  document.body.style.overflow = "hidden";
}

function closeModal() {
  modalExito.classList.remove("active");
  document.body.style.overflow = "";
}

// ==============================================================================
// 13. FUNCIONES AUXILIARES
// ==============================================================================
function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}
