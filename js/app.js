/* Actividades Diariass  */
 
   const DAILY = [
 
    ["06:00", "Mandar a validar y panel de validación"],
    ["06:30", 'Actualizar "EN CURSO"'],
    ["07:00", "CheckList"],
    ["11:00", "Cruce y E5"],
    ["11:30", 'Actualizar el "PANEL DE SSHH"'],
    ["12:00", "Mandar a validar y panel de validación"],
    ["14:00", 'Actualizar el "EN CURSO"'],
    ["17:00", "CheckList Panel"],
    ["17:00", "Mandar a validar y panel de validación"],
    ["20:00", "Panel de SSHH"],
    ["20:10", "Cruce y E5"],
    ["21:30", "Mandar a validar y panel de validación"]
 
];
 
 
/* =====================================================
   Paneles Excepcionales
   DOMINGO=0  LUNES=1  MARTES=2  MIÉRCOLES=3
   JUEVES=4   VIERNES=5  SÁBADO=6
   ===================================================== */
const WEEKLY = {
    1: [
        ["05:00", "TVM"],
        ["07:00", "SMP, SNC, SEMAFORO"],
        ["15:00", "Opersac"],
        ["21:30", "Validación de Patios"]
    ],
    2: [
        ["05:00", "Ingeniería"]
    ],
    4: [
        ["06:30", "SMP"],
        ["21:30", "Validación de Patios"]
    ],
    5: [
        ["05:00", "PDR"]
    ]
};
 
 
const DAYS = ["DOMINGO","LUNES","MARTES","MIÉRCOLES","JUEVES","VIERNES","SÁBADO"];
const MONTHS = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"];
 
/* La tarea se habilita este número de minutos antes de su hora*/
const ENABLE_BEFORE_MIN = 15;
 
/* El checklist se habilita este número de días antes de la fecha */
const CHECKLIST_BEFORE_DAYS = 2;
 
 
/* =====================================================
   GUARDADO
   ===================================================== */
 
let completed = JSON.parse(localStorage.getItem("operaciones_completed") || "{}");
let checklistTarget = localStorage.getItem("checklist_target") || "";
let selectedPeriod = null;
let activities = [];
 
 
/* =====================================================
   FUNCIONES BÁSICAS
   ===================================================== */
 
function pad(n){ return String(n).padStart(2,"0"); }
 
function dateKey(date){
    return `${date.getFullYear()}-${pad(date.getMonth()+1)}-${pad(date.getDate())}`;
}
 
function mins(time){
    const [h,m] = time.split(":").map(Number);
    return h*60 + m;
}
 
function nowMins(date){
    return date.getHours()*60 + date.getMinutes();
}
 
function esc(value){
    return value.replace(/[&<>"']/g, c => ({
        "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"
    }[c]));
}
 
function lastDayOfMonth(year,month){
    return new Date(year, month+1, 0).getDate();
}
 
 
/* =====================================================
   ACTIVIDADES REALIZADAS
   ===================================================== */
 
function activityKey(item,date){
    return `${dateKey(date)}|${item.time}|${item.name}`;
}
 
function isDone(item,date){
    return !!completed[activityKey(item,date)];
}
 
function setDone(item,date,value){
    const key = activityKey(item,date);
    if(value){ completed[key] = true; }
    else{ delete completed[key]; }
    localStorage.setItem("operaciones_completed", JSON.stringify(completed));
}
 
 
/* =====================================================
   CREAR AGENDA DEL DÍA
   ===================================================== */
 
function buildActivities(date){
    return [
        ...DAILY,
        ...(WEEKLY[date.getDay()] || [])
    ]
    .map(item => ({ time:item[0], name:item[1] }))
    .sort((a,b) => mins(a.time) - mins(b.time));
}
 
 
/*Reloj*/
 
function updateClock(){
    const date = new Date();
 
    document.getElementById("currentDate").textContent =
        `${DAYS[date.getDay()]} ${pad(date.getDate())} ${MONTHS[date.getMonth()]}`;
 
    document.getElementById("currentTime").textContent =
        `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}
 
 
/*Tarjeta Principal*/
 
function renderTask(date){
 
    const nowM = nowMins(date);
 
    const pending = activities.find(a => !isDone(a,date));
 
    const time = document.getElementById("taskTime");
    const name = document.getElementById("taskName");
    const button = document.getElementById("taskButton");
    const hint = document.getElementById("taskHint");
 
 
    /* No quedan tareas por hacer */
 
    if(!pending){
        time.textContent = "—";
        name.textContent = "Excelente trabajo, estoy orgullozo de Ti";
        button.disabled = true;
        button.textContent = "SIN PENDIENTES";
        hint.textContent = "✓ Buen descanso";
        return;
    }
 
 
    const start = mins(pending.time);
    const diff = start - nowM;
 
    time.textContent = pending.time;
    name.textContent = pending.name;
 
    button.onclick = () => {
        setDone(pending, date, true);
        renderAll();
    };
 
 
    /* Habilitada: faltan 15 min o menos, o ya llegó / pasó la hora */
 
    if(diff <= ENABLE_BEFORE_MIN){
 
        button.disabled = false;
        button.textContent = "✓ CONFIRMAR";
 
        hint.textContent =
            diff > 0
            ? `Se activa en ${diff} min`
            : "Toca el botón cuando lo hayas hecho";
 
    }
 
    /* Todavía falta */
 
    else{
 
        button.disabled = true;
        button.textContent = "AÚN NO DISPONIBLE";
 
        hint.textContent =
            diff < 60
            ? `Se habilita en ${diff} min`
            : `Se habilita en ${Math.floor(diff/60)} h ${diff%60} min`;
 
    }
 
}
 
 
/* =====================================================
   MODAL: TODAS LAS ACTIVIDADES DEL DÍA
   ===================================================== */
 
function renderTimeline(date){
 
    document.getElementById("todayCount").textContent = activities.length;
 
    document.getElementById("timeline").innerHTML = activities.map((activity,index) => {
 
        const done = isDone(activity,date);
 
        return `
        <div class="timeline-row ${done ? "done" : ""}">
 
            <div class="timeline-time">${activity.time}</div>
 
            <div class="timeline-main">
                <strong>${esc(activity.name)}</strong>
            </div>
 
            <button class="row-check" data-index="${index}">
                ${done ? "✓ HECHO" : "MARCAR"}
            </button>
 
        </div>
        `;
 
    }).join("");
 
    document.querySelectorAll(".row-check").forEach(button => {
        button.onclick = () => {
            const activity = activities[Number(button.dataset.index)];
            setDone(activity, date, !isDone(activity,date));
            renderAll();
        };
    });
 
}
 
 
/* =====================================================
   CHECKLIST — se habilita 2 días antes de quincena o fin de mes
   ===================================================== */
 
function checklistTargetDates(date){
 
    const y = date.getFullYear();
    const m = date.getMonth();
 
    let quincena = new Date(y, m, 15);
    let fin = new Date(y, m, lastDayOfMonth(y,m));
 
    if(quincena < date){
        quincena = new Date(y, m+1, 15);
    }
 
    if(fin < date){
        fin = new Date(y, m+1, lastDayOfMonth(y, m+1));
    }
 
    return quincena <= fin
        ? { date:quincena, type:"quincena" }
        : { date:fin, type:"fin" };
 
}
 
function updateChecklist(date){
 
    const today = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const target = checklistTargetDates(today);
 
    const diffDays = Math.round((target.date - today) / 86400000);
 
    const days = document.getElementById("checklistDays");
    const unit = document.getElementById("checklistUnit");
    const text = document.getElementById("checklistText");
    const button = document.getElementById("programChecklist");
 
    days.textContent = diffDays;
    unit.textContent = diffDays === 1 ? "DÍA" : "DÍAS";
 
    const enabled = diffDays <= CHECKLIST_BEFORE_DAYS;
 
    /* Ya programado para esta fecha */
 
    if(checklistTarget === dateKey(target.date)){
 
        text.textContent = "✓ Ya programado";
        button.disabled = true;
        button.textContent = "PROGRAMADO";
        return;
 
    }
 
    if(enabled){
 
        text.textContent =
            target.type === "quincena"
            ? "para la quincena — puedes programar"
            : "para fin de mes — puedes programar";
 
        button.disabled = false;
        button.textContent = "PROGRAMAR CHECKLIST";
 
    }else{
 
        text.textContent = "para la próxima programación";
        button.disabled = true;
        button.textContent = "AÚN NO DISPONIBLE";
 
    }
 
}
 
 
/* =====================================================
   MODAL: ACTIVIDADES DE HOY
   ===================================================== */
 
const timelineModal = document.getElementById("timelineModal");
 
document.getElementById("openTimeline").onclick = () => {
    timelineModal.classList.add("show");
};
 
document.getElementById("closeTimeline").onclick = () => {
    timelineModal.classList.remove("show");
};
 
 
/* =====================================================
   MODAL: PROGRAMAR CHECKLIST
   ===================================================== */
 
const checklistModal = document.getElementById("checklistModal");
 
document.getElementById("programChecklist").onclick = () => {
 
    if(document.getElementById("programChecklist").disabled) return;
 
    selectedPeriod = null;
 
    document.querySelectorAll(".program-option").forEach(o => o.classList.remove("selected"));
    document.getElementById("saveChecklist").disabled = true;
 
    checklistModal.classList.add("show");
 
};
 
document.getElementById("closeModal").onclick = () => {
    checklistModal.classList.remove("show");
};
 
document.querySelectorAll(".program-option").forEach(option => {
 
    option.onclick = () => {
 
        selectedPeriod = option.dataset.period;
 
        document.querySelectorAll(".program-option").forEach(o => o.classList.remove("selected"));
        option.classList.add("selected");
 
        document.getElementById("saveChecklist").disabled = false;
 
    };
 
});
 
document.getElementById("saveChecklist").onclick = () => {
 
    const date = new Date();
    const y = date.getFullYear();
    const m = date.getMonth();
 
    const target =
        selectedPeriod === "quincena"
        ? new Date(y, m, 15)
        : new Date(y, m, lastDayOfMonth(y,m));
 
    checklistTarget = dateKey(target);
    localStorage.setItem("checklist_target", checklistTarget);
 
    checklistModal.classList.remove("show");
 
    updateChecklist(new Date());
 
};
 
 
/* =====================================================
   RENDER GENERAL
   ===================================================== */
 
function renderAll(){
 
    const date = new Date();
 
    updateClock();
 
    activities = buildActivities(date);
 
    renderTask(date);
    renderTimeline(date);
    updateChecklist(date);
 
}
 
renderAll();
setInterval(renderAll, 1000);
 