const { createApp, ref, computed, reactive, watch } = Vue;

let idCounter = 1;
function generateId(type) {
    return `axl_${type}_${idCounter++}`;
}

const availableComponents = [
    { type: 'panel', label: 'Panel (Container)' },
    { type: 'modal', label: 'Modal / Alert' },
    { type: 'button', label: 'Button' },
    { type: 'input', label: 'Input Field' },
    { type: 'textarea', label: 'Textarea' },
    { type: 'radio', label: 'Radio Button' },
    { type: 'select', label: 'Select Dropdown' },
    { type: 'label', label: 'Label / Text' },
    { type: 'image', label: 'Image' },
    { type: 'table', label: 'Table' },
    { type: 'link', label: 'Hyperlink' },
    { type: 'placeholder', label: 'Placeholder Div' },
    { type: 'timer', label: 'Server Timer (Hidden)' },
    { type: 'rawhtml', label: 'Raw HTML' },
    { type: 'script', label: 'JavaScript' },
    { type: 'style', label: 'CSS Style' }
];

function createComponentInstance(type) {
    const base = {
        id: generateId(type),
        type: type,
        cssClass: '',
        style: '',
        width: '',
        height: '',
        position: 'static',
        top: '',
        bottom: '',
        left: '',
        right: '',
        events: {}, // Server events
        clientEvents: {}, // Client JS events
        reRender: false
    };

    switch (type) {
        case 'panel': return { ...base, children: [], cssClass: 'card' };
        case 'modal': return {
            ...base, title: 'Notice', text: 'This is an alert.', modalType: 'info',
            showBtn1: true, btn1Text: 'OK', btn1Action: 'this.parentNode.parentNode.parentNode.style.display=\'none\';',
            showBtn2: false, btn2Text: 'Cancel', btn2Action: '',
            showBtn3: false, btn3Text: 'Apply', btn3Action: ''
        };
        case 'button': return { ...base, text: 'Click Me', cssClass: 'btn btn-primary', events: { onclick: '// your logic here\n' } };
        case 'input': return { ...base, text: '', inputType: 'text', cssClass: 'prop-input' };
        case 'textarea': return { ...base, text: '', cssClass: 'prop-textarea' };
        case 'radio': return { ...base, text: 'Radio Option' };
        case 'select': return { ...base, options: 'Option 1, Option 2', cssClass: 'prop-input' };
        case 'label': return { ...base, text: 'Label text', reRender: true };
        case 'image': return { ...base, src: 'https://g3pix.com.br/axonasp/apple-icon-60x60.png' };
        case 'table':
            let t = { ...base, rows: 2, cols: 2, cells: {} };
            for (let r = 1; r <= 2; r++) {
                for (let c = 1; c <= 2; c++) {
                    t.cells[`${r}_${c}`] = { id: `${t.id}_${r}_${c}`, type: 'tablecell', text: '', children: [] };
                }
            }
            return t;
        case 'link': return { ...base, text: 'Link Text', src: '#' };
        case 'placeholder': return { ...base, text: 'Content placeholder', cssClass: 'info-banner' };
        case 'timer': return { ...base, delay: 1000, triggerEvent: 'ontimer', events: { ontimer: '// Timer event logic here\n' } };
        case 'rawhtml': return { ...base, text: '<div>Raw HTML</div>' };
        case 'script': return { ...base, text: 'console.log("Hello from AxonLive");' };
        case 'style': return { ...base, text: '/* your css here */' };
    }
    return base;
}

const buildStyleString = (comp) => {
    let s = comp.style || '';
    if (s && !s.endsWith(';')) s += ';';
    if (comp.width) s += `width:${comp.width};`;
    if (comp.height) s += `height:${comp.height};`;
    if (comp.position && comp.position !== 'static') {
        s += `position:${comp.position};`;
        if (comp.top) s += `top:${comp.top};`;
        if (comp.bottom) s += `bottom:${comp.bottom};`;
        if (comp.left) s += `left:${comp.left};`;
        if (comp.right) s += `right:${comp.right};`;
    }
    return s;
};

const ComponentRenderer = {
    name: 'ComponentRenderer',
    props: ['comp', 'selectedId'],
    template: `
        <div :class="['canvas-element', { selected: comp.id === selectedId }]" @click.stop="$emit('select', comp)">
            <div style="font-size:9px; color:#aaa; position:absolute; top:-12px; left:0; background:rgba(255,255,255,0.8); z-index:10; white-space:nowrap; overflow:hidden;">{{ comp.id }}</div>
            
            <div v-if="comp.type === 'panel'" style="min-height: 50px; padding: 5px;" :class="comp.cssClass" :style="computedStyle">
                <component-renderer 
                    v-for="(child, index) in comp.children" 
                    :key="child.id" 
                    :comp="child" 
                    :selected-id="selectedId"
                    @select="$emit('select', $event)"
                    @remove="removeChild(index)">
                </component-renderer>
                <div style="height:20px; background:#f9f9f9; border: 1px dashed #ccc; margin-top:5px; text-align:center; font-size:10px; color:#888;"
                     @dragover.prevent.stop @drop.stop="onDropChild($event, comp)">Drop here to add child</div>
            </div>
            
            <div v-else-if="comp.type === 'modal'" class="window" :class="comp.cssClass" :style="computedStyle">
                <div class="window-header">
                    <span>{{ comp.title }}</span>
                    <span style="cursor:pointer" onclick="alert('Close click preview')">X</span>
                </div>
                <div class="window-body" style="background:#fff;">
                    <div :class="'alert alert-' + comp.modalType" v-if="comp.modalType !== 'none'">
                        {{ comp.text }}
                    </div>
                    <div v-else>
                        {{ comp.text }}
                    </div>
                    <div style="margin-top: 15px; display:flex; justify-content:flex-end; gap:5px;">
                        <button v-if="comp.showBtn1" class="btn btn-primary" disabled>{{ comp.btn1Text }}</button>
                        <button v-if="comp.showBtn2" class="btn btn-secondary" disabled>{{ comp.btn2Text }}</button>
                        <button v-if="comp.showBtn3" class="btn btn-secondary" disabled>{{ comp.btn3Text }}</button>
                    </div>
                </div>
            </div>

            <button v-else-if="comp.type === 'button'" :class="comp.cssClass" :style="computedStyle" disabled>{{ comp.text }}</button>
            <input v-else-if="comp.type === 'input'" :type="comp.inputType" :class="comp.cssClass" :style="computedStyle" :value="comp.text" disabled>
            <textarea v-else-if="comp.type === 'textarea'" :class="comp.cssClass" :style="computedStyle" disabled>{{ comp.text }}</textarea>
            <label v-else-if="comp.type === 'radio'" :style="computedStyle"><input type="radio" disabled> {{ comp.text }}</label>
            <select v-else-if="comp.type === 'select'" :class="comp.cssClass" :style="computedStyle" disabled>
                <option v-for="opt in (comp.options || '').split(',')" :key="opt">{{ opt.trim() }}</option>
            </select>
            <span v-else-if="comp.type === 'label'" :class="comp.cssClass" :style="computedStyle">{{ comp.text }}</span>
            <img v-else-if="comp.type === 'image'" :src="comp.src" :class="comp.cssClass" :style="computedStyle" alt="Image" style="max-width:100%;">
            <a v-else-if="comp.type === 'link'" href="#" :class="comp.cssClass" :style="computedStyle" @click.prevent>{{ comp.text }}</a>
            
            <table v-else-if="comp.type === 'table'" :class="comp.cssClass" :style="computedStyle" border="1" style="border-collapse:collapse; width:100%;">
                <tr v-for="r in comp.rows" :key="r">
                    <td v-for="c in comp.cols" :key="c" style="padding:5px; min-width:40px; min-height:30px; border:1px solid #ccc; vertical-align:top;"
                        :class="{ 'selected-cell': selectedId === comp.id + '_' + r + '_' + c }"
                        @dragover.prevent.stop @drop.stop="onDropChildTable($event, comp, r, c)"
                        @click.stop="$emit('select', getCellData(comp, r, c))">
                        
                        <div style="font-size:10px; color:#666; margin-bottom:2px;" v-if="!getCellData(comp,r,c).text && getCellData(comp,r,c).children.length===0">Cell {{r}}-{{c}}</div>
                        <div v-if="getCellData(comp, r, c).text">{{ getCellData(comp, r, c).text }}</div>
                        
                        <component-renderer 
                            v-for="(child, index) in getCellData(comp, r, c).children" 
                            :key="child.id" 
                            :comp="child" 
                            :selected-id="selectedId"
                            @select="$emit('select', $event)"
                            @remove="removeChildFromCell(comp, r, c, index)">
                        </component-renderer>
                    </td>
                </tr>
            </table>
            
            <div v-else-if="comp.type === 'placeholder'" :class="comp.cssClass" :style="computedStyle" style="background:#eee; padding:10px; text-align:center;">{{ comp.text }}</div>
            <div v-else-if="comp.type === 'timer'" class="timer-component">
                &#9202; Server Timer: {{ comp.id }} ({{ comp.delay }}ms -> {{ comp.triggerEvent }})
            </div>
            <div v-else-if="comp.type === 'rawhtml'" class="canvas-element" style="background:#ffd; border: 1px dashed #aa0; overflow:hidden;">
                <b>Raw HTML:</b> <pre style="font-size:9px; margin:0;">{{ comp.text }}</pre>
            </div>
            <div v-else-if="comp.type === 'script'" class="canvas-element" style="background:#dfd; border: 1px dashed #0a0; overflow:hidden;">
                <b>Script:</b> <pre style="font-size:9px; margin:0;">{{ comp.text }}</pre>
            </div>
            <div v-else-if="comp.type === 'style'" class="canvas-element" style="background:#ddf; border: 1px dashed #00a; overflow:hidden;">
                <b>CSS Style:</b> <pre style="font-size:9px; margin:0;">{{ comp.text }}</pre>
            </div>
        </div>
    `,
    computed: {
        computedStyle() {
            return buildStyleString(this.comp);
        }
    },
    methods: {
        onDropChild(event, parent) {
            const compData = event.dataTransfer.getData('application/json');
            if (compData) {
                try {
                    const parsed = JSON.parse(compData);
                    if (parsed.isNew) {
                        parent.children.push(createComponentInstance(parsed.type));
                    }
                } catch (e) { console.error(e); }
            }
        },
        onDropChildTable(event, tableComp, r, c) {
            const cell = this.getCellData(tableComp, r, c);
            this.onDropChild(event, cell);
        },
        removeChild(index) {
            this.comp.children.splice(index, 1);
        },
        removeChildFromCell(tableComp, r, c, index) {
            const cell = this.getCellData(tableComp, r, c);
            cell.children.splice(index, 1);
        },
        getCellData(tableComp, r, c) {
            const key = r + '_' + c;
            if (!tableComp.cells) tableComp.cells = {};
            if (!tableComp.cells[key]) {
                tableComp.cells[key] = { id: `${tableComp.id}_${r}_${c}`, type: 'tablecell', text: '', children: [] };
            }
            return tableComp.cells[key];
        }
    }
};

const app = createApp({
    components: {
        'component-renderer': ComponentRenderer
    },
    setup() {
        const components = ref([]);
        const selectedComponent = ref(null);
        const showJsonTree = ref(false);
        const newEventName = ref('onclick');
        const newClientEventName = ref('onclick');

        const pageSettings = reactive({
            title: 'AxonLive Application',
            fileName: 'axonlive_app.asp',
            stylesheet: '/css/axonasp.css',
            display: 'block',
            flexDirection: 'row',
            justifyContent: 'flex-start',
            alignItems: 'flex-start'
        });

        watch(() => selectedComponent.value, (val) => {
            if (val && val.type === 'timer') {
                if (!val.events) val.events = {};
                if (!val.events[val.triggerEvent]) {
                    val.events[val.triggerEvent] = '// Timer logic here\n';
                }
            }
        }, { deep: true });

        const onDragStart = (event, comp) => {
            event.dataTransfer.setData('application/json', JSON.stringify({ isNew: true, type: comp.type }));
        };

        const onDrop = (event, parent) => {
            const compData = event.dataTransfer.getData('application/json');
            if (compData) {
                try {
                    const parsed = JSON.parse(compData);
                    if (parsed.isNew) {
                        components.value.push(createComponentInstance(parsed.type));
                    }
                } catch (e) { console.error(e); }
            }
        };

        const selectComponent = (comp) => {
            selectedComponent.value = comp;
        };

        const removeComponent = () => {
            if (selectedComponent.value) {
                let arr = findComponentParentArray(components.value, selectedComponent.value.id);
                if (arr) {
                    let idx = arr.findIndex(c => c.id === selectedComponent.value.id);
                    if (idx >= 0) arr.splice(idx, 1);
                }
                selectedComponent.value = null;
            }
        };

        const addEvent = () => {
            if (selectedComponent.value) {
                if (!selectedComponent.value.events) selectedComponent.value.events = {};
                if (!selectedComponent.value.events[newEventName.value]) {
                    selectedComponent.value.events[newEventName.value] = '// Logic for ' + newEventName.value + '\n';
                }
            }
        };

        const addClientEvent = () => {
            if (selectedComponent.value) {
                if (!selectedComponent.value.clientEvents) selectedComponent.value.clientEvents = {};
                if (!selectedComponent.value.clientEvents[newClientEventName.value]) {
                    selectedComponent.value.clientEvents[newClientEventName.value] = 'alert("clicked");';
                }
            }
        };

        const clearCanvas = () => {
            if (confirm("Are you sure you want to clear the canvas?")) {
                components.value = [];
                selectedComponent.value = null;
                idCounter = 1;
            }
        };

        const findComponentParentArray = (list, id) => {
            for (let i = 0; i < list.length; i++) {
                if (list[i].id === id) return list;
                if (list[i].children) {
                    let res = findComponentParentArray(list[i].children, id);
                    if (res) return res;
                }
                if (list[i].type === 'table' && list[i].cells) {
                    for (const key in list[i].cells) {
                        let cell = list[i].cells[key];
                        if (cell.id === id) return null;
                        if (cell.children) {
                            let res = findComponentParentArray(cell.children, id);
                            if (res) return res;
                        }
                    }
                }
            }
            return null;
        };

        const moveComponent = (direction) => {
            if (!selectedComponent.value) return;
            let arr = findComponentParentArray(components.value, selectedComponent.value.id);
            if (arr) {
                let idx = arr.findIndex(c => c.id === selectedComponent.value.id);
                if (idx >= 0) {
                    if (direction === 'up' && idx > 0) {
                        let temp = arr[idx];
                        arr[idx] = arr[idx - 1];
                        arr[idx - 1] = temp;
                    } else if (direction === 'down' && idx < arr.length - 1) {
                        let temp = arr[idx];
                        arr[idx] = arr[idx + 1];
                        arr[idx + 1] = temp;
                    }
                }
            }
        };

        const generateHTML = (compList, indent = "") => {
            let html = "";
            for (const comp of compList) {
                if (comp.type === 'timer') continue;

                let attrs = `id="${comp.id}"`;
                if (comp.cssClass && comp.type !== 'modal') attrs += ` class="${comp.cssClass}"`;

                const styleStr = buildStyleString(comp);
                if (styleStr) attrs += ` style="${styleStr}"`;

                let hasEvents = false;
                let primaryEvent = "";
                if (comp.events) {
                    for (const evt in comp.events) {
                        hasEvents = true;
                        primaryEvent = evt;
                    }
                }

                if (hasEvents) {
                    const domEvt = primaryEvent.replace(/^on/, '');
                    attrs += ` data-g3al-id="${comp.id}" data-g3al-event="${domEvt}" data-g3al-event-name="${primaryEvent}"`;
                }

                if (comp.clientEvents) {
                    for (const ce in comp.clientEvents) {
                        attrs += ` ${ce}="${comp.clientEvents[ce].replace(/"/g, '&quot;')}"`;
                    }
                }

                if (comp.type === 'panel') {
                    html += `${indent}<div ${attrs}>\n`;
                    html += generateHTML(comp.children, indent + "    ");
                    html += `${indent}</div>\n`;
                } else if (comp.type === 'modal') {
                    let mClass = comp.cssClass ? ` ${comp.cssClass}` : '';
                    html += `${indent}<div ${attrs} class="window${mClass}">\n`;
                    html += `${indent}  <div class="window-header"><span>${comp.title}</span><span style="cursor:pointer" onclick="this.parentNode.parentNode.style.display='none'">X</span></div>\n`;
                    html += `${indent}  <div class="window-body">\n`;
                    if (comp.modalType !== 'none') {
                        html += `${indent}    <div class="alert alert-${comp.modalType}">${comp.text}</div>\n`;
                    } else {
                        html += `${indent}    <div>${comp.text}</div>\n`;
                    }
                    html += `${indent}    <div style="margin-top: 15px; display:flex; justify-content:flex-end; gap:5px;">\n`;
                    if (comp.showBtn1) html += `${indent}      <button class="btn btn-primary" onclick="${comp.btn1Action.replace(/"/g, '&quot;')}">${comp.btn1Text}</button>\n`;
                    if (comp.showBtn2) html += `${indent}      <button class="btn btn-secondary" onclick="${comp.btn2Action.replace(/"/g, '&quot;')}">${comp.btn2Text}</button>\n`;
                    if (comp.showBtn3) html += `${indent}      <button class="btn btn-secondary" onclick="${comp.btn3Action.replace(/"/g, '&quot;')}">${comp.btn3Text}</button>\n`;
                    html += `${indent}    </div>\n`;
                    html += `${indent}  </div>\n`;
                    html += `${indent}</div>\n`;
                } else if (comp.type === 'button') {
                    html += `${indent}<button ${attrs}>${comp.text}</button>\n`;
                } else if (comp.type === 'input') {
                    html += `${indent}<input type="${comp.inputType || 'text'}" ${attrs} value="${comp.text}">\n`;
                } else if (comp.type === 'textarea') {
                    html += `${indent}<textarea ${attrs}>${comp.text}</textarea>\n`;
                } else if (comp.type === 'radio') {
                    html += `${indent}<label ${attrs}><input type="radio" name="${comp.id}_group"> ${comp.text}</label>\n`;
                } else if (comp.type === 'select') {
                    html += `${indent}<select ${attrs}>\n`;
                    const opts = (comp.options || '').split(',');
                    for (const opt of opts) {
                        const o = opt.trim();
                        html += `${indent}  <option value="${o}">${o}</option>\n`;
                    }
                    html += `${indent}</select>\n`;
                } else if (comp.type === 'label') {
                    html += `${indent}<span ${attrs}>${comp.text}</span>\n`;
                } else if (comp.type === 'image') {
                    html += `${indent}<img src="${comp.src}" ${attrs} alt="">\n`;
                } else if (comp.type === 'link') {
                    html += `${indent}<a href="${comp.src || '#'}" ${attrs}>${comp.text}</a>\n`;
                } else if (comp.type === 'table') {
                    html += `${indent}<table ${attrs}>\n`;
                    for (let r = 1; r <= comp.rows; r++) {
                        html += `${indent}  <tr>\n`;
                        for (let c = 1; c <= comp.cols; c++) {
                            let cell = comp.cells[r + '_' + c];
                            html += `${indent}    <td style="padding:5px;">\n`;
                            if (cell) {
                                if (cell.text) html += `${indent}      ${cell.text}\n`;
                                if (cell.children && cell.children.length > 0) {
                                    html += generateHTML(cell.children, indent + "      ");
                                }
                            }
                            html += `${indent}    </td>\n`;
                        }
                        html += `${indent}  </tr>\n`;
                    }
                    html += `${indent}</table>\n`;
                } else if (comp.type === 'placeholder') {
                    html += `${indent}<div ${attrs}>${comp.text}</div>\n`;
                } else if (comp.type === 'rawhtml') {
                    html += `${indent}${comp.text}\n`;
                } else if (comp.type === 'script') {
                    html += `${indent}<script>\n${indent}${comp.text}\n${indent}</script>\n`;
                } else if (comp.type === 'style') {
                    html += `${indent}<style>\n${indent}${comp.text}\n${indent}</style>\n`;
                }
            }
            return html;
        };

        // Collects all reactive components that use a _val variable (reRender=true, stateful types).
        const collectStateComponents = (compList, result) => {
            result = result || [];
            for (const comp of compList) {
                if (comp.reRender && (comp.type === 'label' || comp.type === 'input' || comp.type === 'textarea' || comp.type === 'select')) {
                    result.push(comp);
                }
                if (comp.children) collectStateComponents(comp.children, result);
                if (comp.type === 'table' && comp.cells) {
                    for (const k in comp.cells) {
                        if (comp.cells[k].children) collectStateComponents(comp.cells[k].children, result);
                    }
                }
            }
            return result;
        };

        // Generates the state restore block: reads _val variables from the G3AL server-side store.
        // Without this, _val is undefined on every async re-execution, breaking arithmetic like val + 1.
        const generateStateRestore = (compList) => {
            const stateComps = collectStateComponents(compList, []);
            if (stateComps.length === 0) return "";
            let js = "// Restore persisted component state (survives across async re-executions)\n";
            for (const comp of stateComps) {
                const defaultVal = (comp.text || "").replace(/"/g, '\\"');
                js += `var ${comp.id}_val = AxonLive.GetComponentProperty(sessionID, "${comp.id}", "val");\n`;
                js += `if (${comp.id}_val === null || ${comp.id}_val === "") { ${comp.id}_val = "${defaultVal}"; }\n`;
            }
            return js;
        };

        // Generates the state persist block: writes _val variables back to the G3AL store after events.
        const generateStatePersist = (compList) => {
            const stateComps = collectStateComponents(compList, []);
            if (stateComps.length === 0) return "";
            let js = "    // Persist updated state for the next async call\n";
            for (const comp of stateComps) {
                js += `    AxonLive.SetComponentProperty(sessionID, "${comp.id}", "val", String(${comp.id}_val));\n`;
            }
            return js;
        };

        const generateReRenderCalls = (compList) => {
            let js = "";
            for (const comp of compList) {
                if (comp.reRender && comp.type !== 'timer' && comp.type !== 'script' && comp.type !== 'style' && comp.type !== 'rawhtml') {
                    let attrs = `id="${comp.id}"`;
                    if (comp.cssClass) attrs += ` class="${comp.cssClass}"`;
                    const styleStr = buildStyleString(comp);
                    if (styleStr) attrs += ` style="${styleStr}"`;

                    let inner = "";
                    let tag = "div";
                    // Use explicit null/empty check instead of falsy || to avoid hiding 0 or false values
                    if (comp.type === 'label') { tag = "span"; inner = `'+(${comp.id}_val !== null && ${comp.id}_val !== "" ? ${comp.id}_val : "${comp.text}")+'`; }
                    else if (comp.type === 'button') { tag = "button"; inner = comp.text; }
                    else if (comp.type === 'input') { tag = "input"; inner = ""; attrs += ` value="'+(${comp.id}_val !== null && ${comp.id}_val !== "" ? ${comp.id}_val : "${comp.text}")+'" type="${comp.inputType}"`; }
                    else if (comp.type === 'textarea') { tag = "textarea"; inner = `'+(${comp.id}_val !== null && ${comp.id}_val !== "" ? ${comp.id}_val : "${comp.text}")+'`; }
                    else if (comp.type === 'select') { tag = "select"; inner = `'+(${comp.id}_val || "")+'`; }
                    else { inner = comp.text || ""; }

                    let htmlStr = `<${tag} ${attrs}>${inner}</${tag}>`;
                    if (tag === 'input' || tag === 'img') {
                        htmlStr = `<${tag} ${attrs}>`;
                    }
                    js += `    AxonLive.RegisterComponent("${comp.id}", '${htmlStr}');\n`;
                }
                if (comp.children) {
                    js += generateReRenderCalls(comp.children);
                }
                if (comp.type === 'table' && comp.cells) {
                    for (const k in comp.cells) {
                        if (comp.cells[k].children) js += generateReRenderCalls(comp.cells[k].children);
                    }
                }
            }
            return js;
        };

        const getTimers = (compList) => {
            let timers = [];
            for (const comp of compList) {
                if (comp.type === 'timer') timers.push(comp);
                if (comp.children) timers = timers.concat(getTimers(comp.children));
                if (comp.type === 'table' && comp.cells) {
                    for (const k in comp.cells) {
                        if (comp.cells[k].children) timers = timers.concat(getTimers(comp.cells[k].children));
                    }
                }
            }
            return timers;
        };

        const generateEventSwitch = (compList) => {
            let js = "";
            for (const comp of compList) {
                if (comp.events && Object.keys(comp.events).length > 0) {
                    js += `        case "${comp.id}":\n`;
                    for (const evt in comp.events) {
                        js += `            if (evtName === "${evt}") {\n`;
                        const lines = comp.events[evt].split('\n').map(l => `                ${l}`).join('\n');
                        js += `${lines}`;
                        if (!js.endsWith('\n')) js += '\n';
                        js += `            }\n`;
                    }
                    js += `            break;\n`;
                }
                if (comp.children) {
                    js += generateEventSwitch(comp.children);
                }
                if (comp.type === 'table' && comp.cells) {
                    for (const k in comp.cells) {
                        if (comp.cells[k].children) js += generateEventSwitch(comp.cells[k].children);
                    }
                }
            }
            return js;
        };

        const generatedCode = computed(() => {
            const timers = getTimers(components.value);
            let timerInitCode = "";
            for (const t of timers) {
                timerInitCode += `    // Initialize timer: ${t.id}\n`;
                timerInitCode += `    AxonLive.SetTimer("${t.id}", "${t.triggerEvent}", ${t.delay});\n`;
            }

            const switchLogic = generateEventSwitch(components.value);
            const renderLogic = generateReRenderCalls(components.value);
            const stateRestore = generateStateRestore(components.value);
            const statePersist = generateStatePersist(components.value);
            const htmlLayout = generateHTML(components.value, "    ");

            let mainContainerStyle = ``;
            if (pageSettings.display === 'flex') {
                mainContainerStyle = ` style="display:flex; flex-direction:${pageSettings.flexDirection}; justify-content:${pageSettings.justifyContent}; align-items:${pageSettings.alignItems}; width:100%; height:100%;"`;
            }

            return `<%@ Language="JavaScript" %>
<%
/* Auto-generated by G3pix AxonLive Visual Builder */
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

var sessionID = Session.SessionID;

${stateRestore}
if (AxonLive.IsAsyncRequest) {
    var compID  = AxonLive.EventComponentID;
    var evtName = AxonLive.EventName;

    switch (compID) {
${switchLogic}
    }

${statePersist}
${timerInitCode}
${renderLogic}
    AxonLive.EndAsyncResponse();
}
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${pageSettings.title}</title>
    <link rel="stylesheet" href="${pageSettings.stylesheet}">
</head>
<body>

<div id="main-container">
    <div id="content"${mainContainerStyle}>
${htmlLayout}
    </div>
</div>
<script src="/axonlive/g3axonlive.js"></script>
<script>
    G3AxonLive.init('<%=Server.HTMLEncode(Session.SessionID)%>');
</script>
</body>
</html>`;
        });

        const jsonTree = computed(() => {
            return JSON.stringify(components.value, null, 4);
        });

        const copyCode = () => {
            navigator.clipboard.writeText(generatedCode.value).then(() => {
                alert("Code copied to clipboard!");
            });
        };

        const downloadCode = () => {
            const blob = new Blob([generatedCode.value], { type: 'text/plain' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = pageSettings.fileName || 'axonlive_app.asp';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        };

        return {
            availableComponents,
            components,
            selectedComponent,
            pageSettings,
            showJsonTree,
            generatedCode,
            jsonTree,
            newEventName,
            newClientEventName,
            onDragStart,
            onDrop,
            selectComponent,
            removeComponent,
            moveComponent,
            addEvent,
            addClientEvent,
            clearCanvas,
            copyCode,
            downloadCode
        };
    }
});

app.mount('#app');