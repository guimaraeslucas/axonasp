const { createApp, ref, computed, reactive, watch, onMounted, onUnmounted, nextTick } = Vue;

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
    { type: 'checkbox', label: 'Checkbox' },
    { type: 'checkboxlist', label: 'CheckBox List' },
    { type: 'radio', label: 'Radio Button' },
    { type: 'radiobuttonlist', label: 'RadioButton List' },
    { type: 'bulletedlist', label: 'Bulleted List' },
    { type: 'select', label: 'Select Dropdown' },
    { type: 'listbox', label: 'ListBox' },
    { type: 'label', label: 'Label / Text' },
    { type: 'hr', label: 'Horizontal Rule' },
    { type: 'fileuploader', label: 'File Uploader' },
    { type: 'mdviewer', label: 'MD Viewer' },
    { type: 'hiddenfield', label: 'Hidden Field' },
    { type: 'image', label: 'Image' },
    { type: 'iframe', label: 'iFrame' },
    { type: 'table', label: 'Table' },
    { type: 'link', label: 'Hyperlink' },
    { type: 'placeholder', label: 'Placeholder' },
    { type: 'timer', label: 'Server Timer' },
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
        position: 'absolute',
        top: '10px',
        left: '10px',
        zIndex: '',
        fontFamily: '',
        fontSize: '',
        fontWeight: '',
        fontStyle: '',
        textDecoration: '',
        color: '',
        backgroundColor: '',
        border: '',
        // Flex Item Props
        flexGrow: '',
        flexShrink: '',
        flexBasis: '',
        alignSelf: '',
        order: '',
        events: {}, // Server events
        clientEvents: {}, // Client JS events
        reRender: false
    };

    switch (type) {
        case 'panel': return { ...base, children: [], cssClass: 'card', width: '300px', height: '200px' };
        case 'modal': return {
            ...base, title: 'Notice', text: 'This is an alert.', modalType: 'info',
            width: '400px', height: '200px', showInBuilder: true,
            showBtn1: true, btn1Text: 'OK', btn1Action: `G3AxonLive.closeModal('${base.id}');`,
            showBtn2: false, btn2Text: 'Cancel', btn2Action: `G3AxonLive.closeModal('${base.id}');`,
            showBtn3: false, btn3Text: 'Apply', btn3Action: ''
        };
        case 'button': return { ...base, text: 'Click Me', cssClass: 'btn btn-primary', width: '100px', height: '30px', events: { onclick: '// your logic here\n' } };
        case 'input': return { ...base, text: '', inputType: 'text', cssClass: 'prop-input', width: '150px', height: '25px' };
        case 'textarea': return { ...base, text: '', cssClass: 'prop-textarea', width: '200px', height: '80px' };
        case 'checkbox': return { ...base, text: 'Check Me', checked: false, value: '1' };
        case 'checkboxlist': return { ...base, items: 'Item 1, Item 2, Item 3', width: '200px', height: 'auto', cssClass: '' };
        case 'radio': return { ...base, text: 'Radio Option', name: 'group1', value: '1' };
        case 'radiobuttonlist': return { ...base, items: 'Option 1, Option 2', name: 'rbGroup_' + base.id, width: '200px', height: 'auto' };
        case 'bulletedlist': return { ...base, items: 'List Item 1, List Item 2', listType: 'ul', width: '200px', height: 'auto' };
        case 'select': return { ...base, options: 'Option 1, Option 2', cssClass: 'prop-input', width: '150px', height: '25px' };
        case 'listbox': return { ...base, options: 'Option 1, Option 2, Option 3', multiSelect: false, size: 4, width: '150px', height: '100px' };
        case 'label': return { ...base, text: 'Label text', reRender: true, width: '100px', height: '20px' };
        case 'hr': return { ...base, width: '100%', height: '2px', position: 'static' };
        case 'fileuploader': return {
            ...base, width: '300px', height: '80px', targetDir: '/uploads', maxFileSize: 5242880,
            allowedExtensions: '', blockedExtensions: 'exe,bat', preserveName: true, savedFileName: '',
            showUploadButton: true, uploadButtonText: 'Send File',
            reRender: true // To update result label
        };
        case 'mdviewer': return {
            ...base, width: '400px', height: '300px',
            mdFile: '', // Virtual path to .md file (e.g. /docs/readme.md)
            unsafe: true, // Allow raw HTML in Markdown (G3MD.Unsafe)
            reRender: true
        };
        case 'hiddenfield': return { ...base, value: '', position: 'static', width: '0px', height: '0px' };
        case 'image': return {
            ...base,
            src: 'https://g3pix.com.br/axonasp/apple-icon-60x60.png',
            alt: 'Image',
            title: '',
            width: '60px',
            height: '60px'
        };
        case 'iframe': return { ...base, src: 'https://g3pix.com.br', width: '400px', height: '300px' };
        case 'table':
            let t = { ...base, rows: 2, cols: 2, cells: {}, width: '100%', height: 'auto', showHeader: false, showFooter: false };
            for (let r = 1; r <= 2; r++) {
                for (let c = 1; c <= 2; c++) {
                    t.cells[`${r}_${c}`] = { id: `${t.id}_${r}_${c}`, type: 'tablecell', text: '', children: [] };
                }
            }
            return t;
        case 'link': return { ...base, text: 'Link Text', src: '#', width: '100px', height: '20px' };
        case 'placeholder': return { ...base, text: 'Content placeholder', cssClass: 'info-banner', width: '100%', height: '50px' };
        case 'timer': return { ...base, delay: 1000, triggerEvent: 'ontimer', events: { ontimer: '// Timer event logic here\n' } };
        case 'rawhtml': return { ...base, text: '<div>Raw HTML</div>', width: '200px', height: '100px' };
        case 'script': return { ...base, text: 'console.log("Hello from AxonLive");', serverSide: false };
        case 'style': return { ...base, text: '/* your css here */' };
    }
    return base;
}

const buildStyleString = (comp) => {
    let s = comp.style || '';
    if (s && !s.endsWith(';')) s += ';';
    if (comp.width) s += `width:${comp.width};`;
    if (comp.height) s += `height:${comp.height};`;
    if (comp.zIndex) s += `z-index:${comp.zIndex};`;
    if (comp.fontFamily) s += `font-family:${comp.fontFamily};`;
    if (comp.fontSize) s += `font-size:${comp.fontSize};`;
    if (comp.fontWeight) s += `font-weight:${comp.fontWeight};`;
    if (comp.fontStyle) s += `font-style:${comp.fontStyle};`;
    if (comp.textDecoration) s += `text-decoration:${comp.textDecoration};`;
    if (comp.color) s += `color:${comp.color};`;
    if (comp.backgroundColor) s += `background-color:${comp.backgroundColor};`;
    if (comp.border) s += `border:${comp.border};`;

    // Flex Item Props
    if (comp.flexGrow !== '') s += `flex-grow:${comp.flexGrow};`;
    if (comp.flexShrink !== '') s += `flex-shrink:${comp.flexShrink};`;
    if (comp.flexBasis !== '') s += `flex-basis:${comp.flexBasis};`;
    if (comp.alignSelf !== '') s += `align-self:${comp.alignSelf};`;
    if (comp.order !== '') s += `order:${comp.order};`;

    if (comp.position && comp.position !== 'static') {
        s += `position:${comp.position};`;
        if (comp.top) s += `top:${comp.top};`;
        if (comp.bottom) s += `bottom:${comp.bottom};`;
        if (comp.left) s += `left:${comp.left};`;
        if (comp.right) s += `right:${comp.right};`;
    }
    return s;
};

const escapeHtmlAttr = (value) => String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

const escapeHtmlText = (value) => String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

const escapeJsSingleQuotedString = (value) => String(value || '')
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/\r/g, '\\r')
    .replace(/\n/g, '\\n');

const getPrimaryServerEvent = (comp) => {
    if (!comp.events) return '';
    for (const evt in comp.events) {
        return evt;
    }
    return '';
};

const buildComponentAttrs = (comp, options = {}) => {
    let attrs = `id="${escapeHtmlAttr(comp.id)}"`;
    if (comp.cssClass && !options.skipClass) attrs += ` class="${escapeHtmlAttr(comp.cssClass)}"`;

    const styleStr = buildStyleString(comp);
    if (styleStr && !options.skipStyle) attrs += ` style="${escapeHtmlAttr(styleStr)}"`;

    if (options.includeRuntimeBindings) {
        const primaryEvent = getPrimaryServerEvent(comp);
        if (primaryEvent) {
            const domEvt = primaryEvent.replace(/^on/, '');
            attrs += ` data-g3al-id="${escapeHtmlAttr(comp.id)}" data-g3al-event="${escapeHtmlAttr(domEvt)}" data-g3al-event-name="${escapeHtmlAttr(primaryEvent)}"`;
        }

        if (comp.clientEvents) {
            for (const clientEventName in comp.clientEvents) {
                attrs += ` ${clientEventName}="${escapeHtmlAttr(comp.clientEvents[clientEventName])}"`;
            }
        }
    }

    return { attrs, styleStr };
};

const buildUploaderAction = (compId, fileInputId) =>
    `G3AxonLive.uploadFile(&quot;${escapeHtmlAttr(compId)}&quot;, &quot;${escapeHtmlAttr(fileInputId)}&quot;, &quot;onupload&quot;)`;

const renderableAsyncTypes = {
    modal: true,
    button: true,
    input: true,
    textarea: true,
    checkbox: true,
    checkboxlist: true,
    radio: true,
    radiobuttonlist: true,
    select: true,
    listbox: true,
    label: true,
    hiddenfield: true,
    fileuploader: true,
    mdviewer: true,
    placeholder: true,
    table: true
};

const isRenderableAsyncType = (type) => !!renderableAsyncTypes[type];

const syncTableCells = (compList) => {
    if (!compList) return;
    for (const comp of compList) {
        if (comp.type === 'table') {
            if (!comp.cells) comp.cells = {};
            // Sync header cells
            for (let c = 1; c <= comp.cols; c++) {
                const hKey = `h1_${c}`;
                if (!comp.cells[hKey]) {
                    comp.cells[hKey] = { id: `${comp.id}_${hKey}`, type: 'tablecell', text: '', children: [] };
                }
            }
            // Sync body cells
            for (let r = 1; r <= comp.rows; r++) {
                for (let c = 1; c <= comp.cols; c++) {
                    const key = `${r}_${c}`;
                    if (!comp.cells[key]) {
                        comp.cells[key] = { id: `${comp.id}_${key}`, type: 'tablecell', text: '', children: [] };
                    }
                }
            }
            // Sync footer cells
            for (let c = 1; c <= comp.cols; c++) {
                const fKey = `f1_${c}`;
                if (!comp.cells[fKey]) {
                    comp.cells[fKey] = { id: `${comp.id}_${fKey}`, type: 'tablecell', text: '', children: [] };
                }
            }
        }
        if (comp.children) {
            syncTableCells(comp.children);
        }
        if (comp.type === 'table' && comp.cells) {
            for (const key in comp.cells) {
                if (comp.cells[key].children) {
                    syncTableCells(comp.cells[key].children);
                }
            }
        }
    }
};

const getRenderFunctionName = (compId) => `render_${String(compId || '').replace(/[^a-zA-Z0-9_]/g, '_')}`;

const collectRenderableComponents = (compList, result) => {
    result = result || [];
    for (const comp of compList) {
        if (comp.reRender && isRenderableAsyncType(comp.type)) {
            result.push(comp);
        }
        if (comp.children) collectRenderableComponents(comp.children, result);
        if (comp.type === 'table' && comp.cells) {
            for (const k in comp.cells) {
                if (comp.cells[k].children) collectRenderableComponents(comp.cells[k].children, result);
            }
        }
    }
    return result;
};

const getStateDefaultValue = (comp) => {
    if (comp.type === 'checkbox') return comp.checked ? 'true' : 'false';
    if (comp.type === 'checkboxlist' || comp.type === 'radiobuttonlist' || comp.type === 'listbox' || comp.type === 'select') return '';
    if (comp.type === 'hiddenfield') return comp.value || '';
    return comp.text || comp.value || '';
};

const renderUploaderInner = (comp, resultMarkup) => {
    const fileInputId = `${comp.id}_file`;
    const resultId = `${comp.id}_result`;
    const uploadAction = buildUploaderAction(comp.id, fileInputId);
    const inputOnChange = comp.showUploadButton ? '' : ` onchange="${uploadAction}"`;
    const uploadButton = comp.showUploadButton
        ? `<button type="button" class="btn btn-primary" style="margin-top:5px;" onclick="${uploadAction}">${escapeHtmlText(comp.uploadButtonText || 'Send File')}</button>`
        : '';

    return `<div class="sidebar-header" style="font-size:10px; margin-bottom:5px;">File Upload</div><input type="file" id="${escapeHtmlAttr(fileInputId)}"${inputOnChange}>${uploadButton}<div id="${escapeHtmlAttr(resultId)}" style="font-size:10px; color:#666; margin-top:5px;">Result: ${resultMarkup}</div>`;
};

const hasComponentType = (compList, componentType) => {
    for (const comp of compList) {
        if (comp.type === componentType) return true;
        if (comp.children && hasComponentType(comp.children, componentType)) return true;
        if (comp.type === 'table' && comp.cells) {
            for (const key in comp.cells) {
                if (comp.cells[key].children && hasComponentType(comp.cells[key].children, componentType)) return true;
            }
        }
    }
    return false;
};

const ComponentRenderer = {
    name: 'ComponentRenderer',
    props: ['comp', 'selectedId'],
    data() {
        return {
            isEditing: false
        };
    },
    template: `
        <div v-show="comp.type !== 'modal' || comp.showInBuilder || comp.id === selectedId"
             :id="comp.id" :class="['canvas-element', { selected: comp.id === selectedId }]" 
             :style="computedStyle"
             @mousedown.stop="onMouseDown"
             @click.stop="$emit('select', comp)"
             @dblclick.stop="startEdit"
             @contextmenu.prevent.stop="$emit('context-menu', $event, comp)">
            <div style="font-size:9px; color:#aaa; position:absolute; top:-12px; left:0; background:rgba(255,255,255,0.8); z-index:10; white-space:nowrap; overflow:hidden; pointer-events:none;">{{ comp.id }}</div>
            
            <!-- Resizing handles (only for selected element) -->
            <template v-if="comp.id === selectedId && comp.type !== 'hr'">
                <div class="resize-handle handle-nw" @mousedown.stop="startResize($event, 'nw')"></div>
                <div class="resize-handle handle-ne" @mousedown.stop="startResize($event, 'ne')"></div>
                <div class="resize-handle handle-sw" @mousedown.stop="startResize($event, 'sw')"></div>
                <div class="resize-handle handle-se" @mousedown.stop="startResize($event, 'se')"></div>
                <div class="resize-handle handle-n" @mousedown.stop="startResize($event, 'n')"></div>
                <div class="resize-handle handle-s" @mousedown.stop="startResize($event, 's')"></div>
                <div class="resize-handle handle-e" @mousedown.stop="startResize($event, 'e')"></div>
                <div class="resize-handle handle-w" @mousedown.stop="startResize($event, 'w')"></div>
            </template>

            <div v-if="comp.type === 'panel'" style="min-height: 50px; padding: 5px; width:100%; height:100%; position:relative;" :class="comp.cssClass">
                <component-renderer 
                    v-for="(child, index) in comp.children" 
                    :key="child.id" 
                    :comp="child" 
                    :selected-id="selectedId"
                    @select="$emit('select', $event)"
                    @drag-start="$emit('drag-start', $event)"
                    @resize-start="$emit('resize-start', $event)"
                    @context-menu="$emit('context-menu', arguments[0], arguments[1])"
                    @remove="removeChild(index)">
                </component-renderer>
                <div v-if="comp.position === 'static'" style="height:20px; background:#f9f9f9; border: 1px dashed #ccc; margin-top:5px; text-align:center; font-size:10px; color:#888;"
                     @dragover.prevent.stop @drop.stop="onDropChild($event, comp)">Drop here to add child</div>
            </div>
            
            <div v-else-if="comp.type === 'modal'" class="window" :class="comp.cssClass" style="width:100%; height:100%;">
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

            <button v-else-if="comp.type === 'button'" :class="comp.cssClass" style="width:100%; height:100%;" disabled>{{ comp.text }}</button>
            <input v-else-if="comp.type === 'input'" :type="comp.inputType" :class="comp.cssClass" style="width:100%; height:100%;" :value="comp.text" disabled>
            <textarea v-else-if="comp.type === 'textarea'" :class="comp.cssClass" style="width:100%; height:100%;" disabled>{{ comp.text }}</textarea>
            
            <label v-else-if="comp.type === 'checkbox'" style="width:100%; height:100%;"><input type="checkbox" :checked="comp.checked" disabled> {{ comp.text }}</label>
            
            <div v-else-if="comp.type === 'checkboxlist'" :class="comp.cssClass" style="width:100%; height:100%;">
                <div v-for="item in (comp.items || '').split(',')" :key="item">
                    <label><input type="checkbox" disabled> {{ item.trim() }}</label>
                </div>
            </div>

            <label v-else-if="comp.type === 'radio'" style="width:100%; height:100%;"><input type="radio" :name="comp.name" disabled> {{ comp.text }}</label>
            
            <div v-else-if="comp.type === 'radiobuttonlist'" style="width:100%; height:100%;">
                <div v-for="item in (comp.items || '').split(',')" :key="item">
                    <label><input type="radio" :name="comp.name" disabled> {{ item.trim() }}</label>
                </div>
            </div>

            <div v-else-if="comp.type === 'bulletedlist'" style="width:100%; height:100%;">
                <component :is="comp.listType">
                    <li v-for="item in (comp.items || '').split(',')" :key="item">{{ item.trim() }}</li>
                </component>
            </div>

            <select v-else-if="comp.type === 'select'" :class="comp.cssClass" style="width:100%; height:100%;" disabled>
                <option v-for="opt in (comp.options || '').split(',')" :key="opt">{{ opt.trim() }}</option>
            </select>

            <select v-else-if="comp.type === 'listbox'" :class="comp.cssClass" :multiple="comp.multiSelect" :size="comp.size" style="width:100%; height:100%;" disabled>
                <option v-for="opt in (comp.options || '').split(',')" :key="opt">{{ opt.trim() }}</option>
            </select>
            
            <span v-else-if="comp.type === 'label'" :class="comp.cssClass" style="width:100%; height:100%;">
                <input v-if="isEditing" type="text" v-model="comp.text" @blur="stopEdit" @keyup.enter="stopEdit" class="prop-input" style="width:100%">
                <template v-else>{{ comp.text }}</template>
            </span>

            <hr v-else-if="comp.type === 'hr'" :style="computedStyle">

            <div v-else-if="comp.type === 'fileuploader'" class="card" style="width:100%; height:100%; padding:10px;">
                <div style="font-weight:bold; margin-bottom:5px;">FILE UPLOADER</div>
                <input type="file" style="width:100%" disabled>
                <button v-if="comp.showUploadButton" type="button" class="btn btn-primary" style="margin-top:5px;" disabled>{{ comp.uploadButtonText || 'Send File' }}</button>
                <div style="font-size:10px; color:#666; margin-top:5px;">Result: Ready to upload</div>
            </div>

            <div v-else-if="comp.type === 'mdviewer'" style="width:100%; height:100%; overflow:auto; border:1px dashed #6699cc; background:#f8faff; padding:8px; box-sizing:border-box;">
                <div style="font-size:10px; font-weight:bold; color:#003399; border-bottom:1px solid #c0c8d8; padding-bottom:3px; margin-bottom:6px;">MD VIEWER</div>
                <div v-if="comp.mdFile" style="font-size:11px; color:#444; font-style:italic;">File: {{ comp.mdFile }}</div>
                <div v-else style="font-size:11px; color:#999; font-style:italic;">No file configured — content will be empty on load.</div>
                <div style="font-size:10px; color:#888; margin-top:4px;">Rendered via G3MD on server side.</div>
            </div>

            <div v-else-if="comp.type === 'hiddenfield'" style="width:20px; height:20px; background:#ddd; border:1px solid #333; text-align:center; font-size:10px;">H</div>
            
            <img v-else-if="comp.type === 'image'" :src="comp.src" :alt="comp.alt || ''" :title="comp.title || ''" :class="comp.cssClass" style="max-width:100%; width:100%; height:100%;">
            
            <iframe v-else-if="comp.type === 'iframe'" :src="comp.src" style="width:100%; height:100%; border:1px solid #ccc;"></iframe>

            <a v-else-if="comp.type === 'link'" href="#" :class="comp.cssClass" style="width:100%; height:100%;" @click.prevent>
                <input v-if="isEditing" type="text" v-model="comp.text" @blur="stopEdit" @keyup.enter="stopEdit" class="prop-input" style="width:100%">
                <template v-else>{{ comp.text }}</template>
            </a>
            
            <table v-else-if="comp.type === 'table'" :class="comp.cssClass" border="1" style="border-collapse:collapse; width:100%; height:100%;">
                <thead v-if="comp.showHeader">
                    <tr>
                        <th v-for="c in comp.cols" :key="'h'+c" style="padding:5px; background:#eee;" @dblclick.stop="startEditCell(comp, 'h', 1, c)">
                            <input v-if="isEditingCell(comp, 'h', 1, c)" type="text" v-model="getCellData(comp, 'h', 1, c).text" @blur="stopEdit" @keyup.enter="stopEdit" class="prop-input" style="width:100%">
                            <template v-else>{{ getCellData(comp, 'h', 1, c).text || 'Header ' + c }}</template>
                        </th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="r in comp.rows" :key="r">
                        <td v-for="c in comp.cols" :key="c" style="padding:5px; min-width:40px; min-height:30px; border:1px solid #ccc; vertical-align:top; position:relative;"
                            :class="{ 'selected-cell': selectedId === comp.id + '_' + r + '_' + c }"
                            @dragover.prevent.stop @drop.stop="onDropChildTable($event, comp, r, c)"
                            @click.stop="$emit('select', getCellData(comp, '', r, c))"
                            @dblclick.stop="startEditCell(comp, '', r, c)"
                            @contextmenu.prevent.stop="$emit('context-menu', $event, getCellData(comp, '', r, c))">
                            
                            <div style="font-size:10px; color:#666; margin-bottom:2px;" v-if="!isEditingCell(comp, '', r, c) && !getCellData(comp, '', r, c).text && getCellData(comp, '', r, c).children.length===0">Cell {{r}}-{{c}}</div>
                            
                            <input v-if="isEditingCell(comp, '', r, c)" type="text" v-model="getCellData(comp, '', r, c).text" @blur="stopEdit" @keyup.enter="stopEdit" class="prop-input" style="width:100%">
                            <div v-else-if="getCellData(comp, '', r, c).text">{{ getCellData(comp, '', r, c).text }}</div>
                            
                            <component-renderer 
                                v-for="(child, index) in getCellData(comp, '', r, c).children" 
                                :key="child.id" 
                                :comp="child" 
                                :selected-id="selectedId"
                                @select="$emit('select', $event)"
                                @drag-start="$emit('drag-start', $event)"
                                @resize-start="$emit('resize-start', $event)"
                                @context-menu="$emit('context-menu', arguments[0], arguments[1])"
                                @remove="removeChildFromCell(comp, r, c, index)">
                            </component-renderer>
                        </td>
                    </tr>
                </tbody>
                <tfoot v-if="comp.showFooter">
                    <tr>
                        <td v-for="c in comp.cols" :key="'f'+c" style="padding:5px; background:#eee;" @dblclick.stop="startEditCell(comp, 'f', 1, c)">
                            <input v-if="isEditingCell(comp, 'f', 1, c)" type="text" v-model="getCellData(comp, 'f', 1, c).text" @blur="stopEdit" @keyup.enter="stopEdit" class="prop-input" style="width:100%">
                            <template v-else>{{ getCellData(comp, 'f', 1, c).text || 'Footer ' + c }}</template>
                        </td>
                    </tr>
                </tfoot>
            </table>
            
            <div v-else-if="comp.type === 'placeholder'" :class="comp.cssClass" style="background:#eee; padding:10px; text-align:center; width:100%; height:100%;">{{ comp.text }}</div>
            
            <div v-else-if="comp.type === 'rawhtml'" style="width:100%; height:100%; border:1px dashed #aa0; background:#ffd; padding:5px; overflow:hidden;">
                <div style="font-size:9px; color:#aa0; border-bottom:1px solid #aa0; margin-bottom:2px;">RAW HTML</div>
                <div v-html="comp.text"></div>
            </div>
        </div>
    `,
    computed: {
        computedStyle() {
            return buildStyleString(this.comp);
        }
    },
    methods: {
        startEdit() {
            if (['label', 'link'].includes(this.comp.type)) {
                this.isEditing = true;
            }
        },
        startEditCell(table, prefix, r, c) {
            this.isEditing = table.id + '_' + prefix + r + '_' + c;
        },
        isEditingCell(table, prefix, r, c) {
            return this.isEditing === table.id + '_' + prefix + r + '_' + c;
        },
        stopEdit() {
            this.isEditing = false;
        },
        onMouseDown(e) {
            this.$emit('select', this.comp);
            this.$emit('drag-start', { event: e, comp: this.comp });
        },
        startResize(e, dir) {
            this.$emit('resize-start', { event: e, comp: this.comp, dir: dir });
        },
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
            const cell = this.getCellData(tableComp, '', r, c);
            this.onDropChild(event, cell);
        },
        removeChild(index) {
            this.comp.children.splice(index, 1);
        },
        removeChildFromCell(tableComp, r, c, index) {
            const cell = this.getCellData(tableComp, '', r, c);
            cell.children.splice(index, 1);
        },
        getCellData(tableComp, prefix, r, c) {
            if (c === undefined) {
                c = r;
                r = prefix;
                prefix = '';
            }
            const key = (prefix || '') + r + '_' + c;
            if (!tableComp.cells) tableComp.cells = {};
            if (!tableComp.cells[key]) {
                tableComp.cells[key] = { id: `${tableComp.id}_${key}`, type: 'tablecell', text: '', children: [] };
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
        const allComponents = ref([]);
        const selectedComponent = ref(null);
        const showJsonTree = ref(false);
        const codePaneState = ref('normal'); // 'normal', 'maximized', 'hidden'
        const newEventName = ref('onclick');
        const newClientEventName = ref('onclick');
        const editableJsonTree = ref('');
        const aspCodeEl = ref(null);
        const fileInput = ref(null);

        // UNDO / REDO / COPY / PASTE
        const history = ref([]);
        const historyIndex = ref(-1);
        const clipboard = ref(null);

        const contextMenu = reactive({ show: false, x: 0, y: 0 });

        const saveHistory = () => {
            if (historyIndex.value < history.value.length - 1) {
                history.value = history.value.slice(0, historyIndex.value + 1);
            }
            history.value.push(JSON.stringify(allComponents.value));
            if (history.value.length > 50) history.value.shift();
            else historyIndex.value++;
        };

        const undo = () => {
            if (historyIndex.value > 0) {
                historyIndex.value--;
                allComponents.value = JSON.parse(history.value[historyIndex.value]);
                contextMenu.show = false;
            }
        };

        const redo = () => {
            if (historyIndex.value < history.value.length - 1) {
                historyIndex.value++;
                allComponents.value = JSON.parse(history.value[historyIndex.value]);
                contextMenu.show = false;
            }
        };

        const copyComp = () => {
            if (selectedComponent.value) {
                clipboard.value = JSON.parse(JSON.stringify(selectedComponent.value));
            }
            contextMenu.show = false;
        };

        function regenerateIds(comp) {
            comp.id = generateId(comp.type);
            if (comp.children) comp.children.forEach(regenerateIds);
            if (comp.cells) {
                for (let k in comp.cells) regenerateIds(comp.cells[k]);
            }
        }

        const pasteComp = () => {
            if (clipboard.value) {
                const newComp = JSON.parse(JSON.stringify(clipboard.value));
                regenerateIds(newComp);

                if (newComp.position !== 'static') {
                    newComp.top = ((parseInt(newComp.top) || 0) + 20) + 'px';
                    newComp.left = ((parseInt(newComp.left) || 0) + 20) + 'px';
                }

                if (selectedComponent.value && selectedComponent.value.type === 'panel') {
                    selectedComponent.value.children.push(newComp);
                } else if (selectedComponent.value && selectedComponent.value.type === 'tablecell') {
                    selectedComponent.value.children.push(newComp);
                } else {
                    allComponents.value.push(newComp);
                }
                saveHistory();
            }
            contextMenu.show = false;
        };

        const duplicateComp = () => {
            copyComp();
            pasteComp();
            contextMenu.show = false;
        };

        const onContextMenu = (e, comp) => {
            if (comp) selectComponent(comp);
            contextMenu.x = e.clientX;
            contextMenu.y = e.clientY;
            contextMenu.show = true;
        };

        const toggleCodePane = () => {
            if (codePaneState.value === 'normal') codePaneState.value = 'maximized';
            else if (codePaneState.value === 'maximized') codePaneState.value = 'hidden';
            else codePaneState.value = 'normal';
        };

        // File Save/Load
        const saveModel = () => {
            const blob = new Blob([JSON.stringify(allComponents.value, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = (pageSettings.fileName || 'model').replace('.asp', '') + '.g3al';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        };

        const openModel = () => {
            fileInput.value.click();
        };

        const handleFileOpen = (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (ev) => {
                try {
                    allComponents.value = JSON.parse(ev.target.result);
                    saveHistory();
                } catch (err) {
                    alert("Invalid .g3al file format");
                }
            };
            reader.readAsText(file);
            e.target.value = '';
        };

        // Compute separated components
        const components = computed(() => {
            return allComponents.value.filter(c => !['timer', 'script', 'style', 'hiddenfield'].includes(c.type));
        });

        const functionalComponents = computed(() => {
            return allComponents.value.filter(c => ['timer', 'script', 'style', 'hiddenfield'].includes(c.type));
        });

        // DRAG & RESIZE STATE
        const dragData = reactive({
            isDragging: false,
            isResizing: false,
            activeComp: null,
            startX: 0,
            startY: 0,
            startTop: 0,
            startLeft: 0,
            startWidth: 0,
            startHeight: 0,
            resizeDir: ''
        });

        const pageSettings = reactive({
            title: 'AxonLive Application',
            fileName: 'axonlive_app.asp',
            stylesheet: '/css/axonasp.css',
            display: 'block',
            flexDirection: 'row',
            justifyContent: 'flex-start',
            alignItems: 'flex-start',
            canvasWidth: 800,
            canvasHeight: 400
        });

        const handleDragStart = ({ event, comp }) => {
            dragData.isDragging = true;
            dragData.activeComp = comp;
            dragData.startX = event.clientX;
            dragData.startY = event.clientY;
            dragData.startTop = parseInt(comp.top) || 0;
            dragData.startLeft = parseInt(comp.left) || 0;
            contextMenu.show = false;
        };

        const handleResizeStart = ({ event, comp, dir }) => {
            dragData.isResizing = true;
            dragData.activeComp = comp;
            dragData.resizeDir = dir;
            dragData.startX = event.clientX;
            dragData.startY = event.clientY;
            contextMenu.show = false;

            const el = document.getElementById(comp.id);
            if (el) {
                dragData.startWidth = el.offsetWidth;
                dragData.startHeight = el.offsetHeight;
            } else {
                dragData.startWidth = parseInt(comp.width) || 100;
                dragData.startHeight = parseInt(comp.height) || 50;
            }
            dragData.startTop = parseInt(comp.top) || 0;
            dragData.startLeft = parseInt(comp.left) || 0;
        };

        const onMouseMove = (e) => {
            if (dragData.isDragging && dragData.activeComp) {
                const dx = e.clientX - dragData.startX;
                const dy = e.clientY - dragData.startY;

                if (dragData.activeComp.position !== 'static') {
                    const newTop = (dragData.startTop + dy);
                    const newLeft = (dragData.startLeft + dx);

                    dragData.activeComp.top = newTop + 'px';
                    dragData.activeComp.left = newLeft + 'px';

                    // Auto-extend canvas if dragging outside
                    if (newLeft + 100 > pageSettings.canvasWidth) pageSettings.canvasWidth = newLeft + 200;
                    if (newTop + 50 > pageSettings.canvasHeight) pageSettings.canvasHeight = newTop + 100;
                }
            } else if (dragData.isResizing && dragData.activeComp) {
                const dx = e.clientX - dragData.startX;
                const dy = e.clientY - dragData.startY;
                const dir = dragData.resizeDir;

                let newWidth = dragData.startWidth;
                let newHeight = dragData.startHeight;
                let newTop = dragData.startTop;
                let newLeft = dragData.startLeft;

                if (dir.includes('e')) newWidth = dragData.startWidth + dx;
                if (dir.includes('s')) newHeight = dragData.startHeight + dy;
                if (dir.includes('w')) {
                    newWidth = dragData.startWidth - dx;
                    newLeft = dragData.startLeft + dx;
                }
                if (dir.includes('n')) {
                    newHeight = dragData.startHeight - dy;
                    newTop = dragData.startTop + dy;
                }

                if (newWidth > 10) {
                    dragData.activeComp.width = newWidth + 'px';
                    if (dir.includes('w') && dragData.activeComp.position !== 'static') dragData.activeComp.left = newLeft + 'px';

                    // Auto-extend width
                    const rightEdge = (parseInt(dragData.activeComp.left) || 0) + newWidth;
                    if (rightEdge > pageSettings.canvasWidth) pageSettings.canvasWidth = rightEdge + 50;
                }
                if (newHeight > 10) {
                    dragData.activeComp.height = newHeight + 'px';
                    if (dir.includes('n') && dragData.activeComp.position !== 'static') dragData.activeComp.top = newTop + 'px';

                    // Auto-extend height
                    const bottomEdge = (parseInt(dragData.activeComp.top) || 0) + newHeight;
                    if (bottomEdge > pageSettings.canvasHeight) pageSettings.canvasHeight = bottomEdge + 50;
                }
            }
        };

        const onMouseUp = () => {
            if (dragData.isDragging || dragData.isResizing) {
                saveHistory();
            }
            dragData.isDragging = false;
            dragData.isResizing = false;
            dragData.activeComp = null;
        };

        onMounted(() => {
            window.addEventListener('mousemove', onMouseMove);
            window.addEventListener('mouseup', onMouseUp);
            window.addEventListener('click', (e) => {
                if (!e.target.closest('.context-menu')) contextMenu.show = false;
            });
            saveHistory(); // Initial state
            highlightGeneratedCode();
        });

        onUnmounted(() => {
            window.removeEventListener('mousemove', onMouseMove);
            window.removeEventListener('mouseup', onMouseUp);
        });

        watch(() => allComponents.value, (val) => {
            syncTableCells(val);
            if (!dragData.isDragging && !dragData.isResizing) {
                editableJsonTree.value = JSON.stringify(val, null, 4);
            }
        }, { deep: true });

        const updateFromJson = () => {
            try {
                const parsed = JSON.parse(editableJsonTree.value);
                if (Array.isArray(parsed)) {
                    allComponents.value = parsed;
                    saveHistory();
                }
            } catch (e) { }
        };

        const highlightGeneratedCode = () => {
            if (showJsonTree.value) return;
            nextTick(() => {
                if (aspCodeEl.value && window.hljs && typeof window.hljs.highlightElement === 'function') {
                    window.hljs.highlightElement(aspCodeEl.value);
                }
            });
        };

        watch(showJsonTree, (isJsonVisible) => {
            if (!isJsonVisible) {
                highlightGeneratedCode();
            }
        });

        const onDragStart = (event, comp) => {
            event.dataTransfer.setData('application/json', JSON.stringify({ isNew: true, type: comp.type }));
            contextMenu.show = false;
        };

        const onDrop = (event, parent) => {
            const compData = event.dataTransfer.getData('application/json');
            if (compData) {
                try {
                    const parsed = JSON.parse(compData);
                    if (parsed.isNew) {
                        const newComp = createComponentInstance(parsed.type);
                        if (['timer', 'script', 'style'].includes(parsed.type)) {
                            allComponents.value.push(newComp);
                        } else {
                            allComponents.value.push(newComp);
                        }
                        selectComponent(newComp);
                        saveHistory();
                    }
                } catch (e) { console.error(e); }
            }
        };

        const selectComponent = (comp) => {
            selectedComponent.value = comp;
        };

        const removeComponent = () => {
            if (selectedComponent.value) {
                let arr = findComponentParentArray(allComponents.value, selectedComponent.value.id);
                if (arr) {
                    let idx = arr.findIndex(c => c.id === selectedComponent.value.id);
                    if (idx >= 0) arr.splice(idx, 1);
                }
                selectedComponent.value = null;
                saveHistory();
            }
            contextMenu.show = false;
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
                allComponents.value = [];
                selectedComponent.value = null;
                idCounter = 1;
                pageSettings.canvasWidth = 800;
                pageSettings.canvasHeight = 400;
                saveHistory();
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
            let arr = findComponentParentArray(allComponents.value, selectedComponent.value.id);
            if (arr) {
                let idx = arr.findIndex(c => c.id === selectedComponent.value.id);
                if (idx >= 0) {
                    if (direction === 'up' && idx > 0) {
                        let temp = arr[idx];
                        arr[idx] = arr[idx - 1];
                        arr[idx - 1] = temp;
                        saveHistory();
                    } else if (direction === 'down' && idx < arr.length - 1) {
                        let temp = arr[idx];
                        arr[idx] = arr[idx + 1];
                        arr[idx + 1] = temp;
                        saveHistory();
                    }
                }
            }
        };

        const generateHTML = (compList, indent = "") => {
            syncTableCells(compList);
            let html = "";
            for (const comp of compList) {
                if (['timer', 'script', 'style'].includes(comp.type)) continue;

                if (comp.reRender && isRenderableAsyncType(comp.type)) {
                    const fn = getRenderFunctionName(comp.id);
                    html += `${indent}<%=${fn}()%>\n`;
                    continue;
                }

                const { attrs, styleStr } = buildComponentAttrs(comp, { includeRuntimeBindings: true, skipClass: comp.type === 'modal', skipStyle: comp.type === 'modal' });

                if (comp.type === 'panel') {
                    html += `${indent}<div ${attrs}>\n`;
                    html += generateHTML(comp.children, indent + "    ");
                    html += `${indent}</div>\n`;
                } else if (comp.type === 'modal') {
                    let mClass = comp.cssClass ? ` ${comp.cssClass}` : '';
                    html += `${indent}<div ${attrs} class="window${escapeHtmlAttr(mClass)}" style="${escapeHtmlAttr(`display:none; ${styleStr}`)}">\n`;
                    html += `${indent}  <div class="window-header"><span>${comp.title}</span><span style="cursor:pointer" onclick="G3AxonLive.closeModal('${comp.id}')">X</span></div>\n`;
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
                } else if (comp.type === 'checkbox') {
                    const chk = comp.checked ? ' checked' : '';
                    html += `${indent}<label ${attrs}><input type="checkbox"${chk} value="${comp.value || '1'}"> ${comp.text}</label>\n`;
                } else if (comp.type === 'checkboxlist') {
                    html += `${indent}<div ${attrs} data-g3al-type="checkboxlist">\n`;
                    const items = (comp.items || '').split(',');
                    for (const item of items) {
                        const itm = item.trim();
                        html += `${indent}    <label><input type="checkbox" value="${itm}"> ${itm}</label><br>\n`;
                    }
                    html += `${indent}</div>\n`;
                } else if (comp.type === 'radio') {
                    html += `${indent}<label ${attrs}><input type="radio" name="${comp.name}" value="${comp.value || '1'}"> ${comp.text}</label>\n`;
                } else if (comp.type === 'radiobuttonlist') {
                    html += `${indent}<div ${attrs} data-g3al-type="radiobuttonlist">\n`;
                    const items = (comp.items || '').split(',');
                    for (const item of items) {
                        const itm = item.trim();
                        html += `${indent}    <label><input type="radio" name="${comp.name}" value="${itm}"> ${itm}</label><br>\n`;
                    }
                    html += `${indent}</div>\n`;
                } else if (comp.type === 'bulletedlist') {
                    html += `${indent}<${comp.listType} ${attrs}>\n`;
                    const items = (comp.items || '').split(',');
                    for (const item of items) {
                        const itm = item.trim();
                        html += `${indent}    <li>${itm}</li>\n`;
                    }
                    html += `${indent}</${comp.listType}>\n`;
                } else if (comp.type === 'select') {
                    html += `${indent}<select ${attrs}>\n`;
                    const opts = (comp.options || '').split(',');
                    for (const opt of opts) {
                        const o = opt.trim();
                        html += `${indent}  <option value="${o}">${o}</option>\n`;
                    }
                    html += `${indent}</select>\n`;
                } else if (comp.type === 'listbox') {
                    const multi = comp.multiSelect ? ' multiple' : '';
                    const size = comp.size ? ` size="${comp.size}"` : '';
                    html += `${indent}<select ${attrs}${multi}${size}>\n`;
                    const opts = (comp.options || '').split(',');
                    for (const opt of opts) {
                        const o = opt.trim();
                        html += `${indent}  <option value="${o}">${o}</option>\n`;
                    }
                    html += `${indent}</select>\n`;
                } else if (comp.type === 'label') {
                    html += `${indent}<span ${attrs}>${comp.text}</span>\n`;
                } else if (comp.type === 'hr') {
                    html += `${indent}<hr ${attrs}>\n`;
                } else if (comp.type === 'fileuploader') {
                    html += `${indent}<div ${attrs}>\n`;
                    html += `${indent}    ${renderUploaderInner(comp, `<%=Server.HTMLEncode(String(${comp.id}_val || "Ready"))%>`)}\n`;
                    html += `${indent}</div>\n`;
                } else if (comp.type === 'mdviewer') {
                    // _val holds the G3MD-rendered HTML; emit raw (do not HTMLEncode)
                    html += `${indent}<div ${attrs}><%=${comp.id}_val%></div>\n`;
                } else if (comp.type === 'hiddenfield') {
                    html += `${indent}<input type="hidden" ${attrs} value="${comp.value || ''}">\n`;
                } else if (comp.type === 'image') {
                    html += `${indent}<img src="${escapeHtmlAttr(comp.src || '')}" ${attrs} alt="${escapeHtmlAttr(comp.alt || '')}" title="${escapeHtmlAttr(comp.title || '')}">\n`;
                } else if (comp.type === 'iframe') {
                    html += `${indent}<iframe src="${comp.src}" ${attrs} frameborder="0"></iframe>\n`;
                } else if (comp.type === 'link') {
                    html += `${indent}<a href="${comp.src || '#'}" ${attrs}>${comp.text}</a>\n`;
                } else if (comp.type === 'table') {
                    html += `${indent}<table ${attrs}>\n`;
                    if (comp.showHeader) {
                        html += `${indent}  <thead>\n${indent}    <tr>\n`;
                        for (let c = 1; c <= comp.cols; c++) {
                            let cell = comp.cells['h1_' + c];
                            html += `${indent}      <th>${(cell && cell.text) || 'Header ' + c}</th>\n`;
                        }
                        html += `${indent}    </tr>\n${indent}  </thead>\n`;
                    }
                    html += `${indent}  <tbody>\n`;
                    for (let r = 1; r <= comp.rows; r++) {
                        html += `${indent}    <tr>\n`;
                        for (let c = 1; c <= comp.cols; c++) {
                            let cell = comp.cells[r + '_' + c];
                            html += `${indent}      <td style="padding:5px;">\n`;
                            if (cell) {
                                if (cell.text) html += `${indent}        ${cell.text}\n`;
                                if (cell.children && cell.children.length > 0) {
                                    html += generateHTML(cell.children, indent + "        ");
                                }
                            }
                            html += `${indent}      </td>\n`;
                        }
                        html += `${indent}    </tr>\n`;
                    }
                    html += `${indent}  </tbody>\n`;
                    if (comp.showFooter) {
                        html += `${indent}  <tfoot>\n${indent}    <tr>\n`;
                        for (let c = 1; c <= comp.cols; c++) {
                            let cell = comp.cells['f1_' + c];
                            html += `${indent}      <td>${(cell && cell.text) || 'Footer ' + c}</td>\n`;
                        }
                        html += `${indent}    </tr>\n${indent}  </tfoot>\n`;
                    }
                    html += `${indent}</table>\n`;
                } else if (comp.type === 'placeholder') {
                    html += `${indent}<div ${attrs}>${comp.text}</div>\n`;
                } else if (comp.type === 'rawhtml') {
                    html += `${indent}${comp.text}\n`;
                }
            }
            return html;
        };

        const collectStateComponents = (compList, result) => {
            result = result || [];
            for (const comp of compList) {
                if (comp.reRender && isRenderableAsyncType(comp.type)) {
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

        const generateStateRestore = (compList) => {
            const stateComps = collectStateComponents(compList, []);
            if (stateComps.length === 0) return "";
            let js = "// Restore persisted component state (survives across async re-executions)\n";
            for (const comp of stateComps) {
                js += `var ${comp.id}_val = AxonLive.GetComponentProperty(sessionID, "${comp.id}", "val");\n`;
                if (comp.type === 'mdviewer') {
                    // MD Viewer: on first load, read and render the configured markdown file via G3FILES + G3MD.
                    // Server.MapPath resolves the virtual path so both HTTP server and CLI work correctly.
                    // Use falsy check: GetComponentProperty returns undefined (not null/"") on first load
                    js += `if (!${comp.id}_val) {\n`;
                    if (comp.mdFile) {
                        const safePath = comp.mdFile.replace(/\\/g, '/').replace(/"/g, '\\"');
                        js += `    var __mdpath_${comp.id} = Server.MapPath("${safePath}");\n`;
                        js += `    var __g3files_${comp.id} = Server.CreateObject("G3FILES");\n`;
                        js += `    if (__g3files_${comp.id}.Exists(__mdpath_${comp.id})) {\n`;
                        js += `        var __g3md_${comp.id} = Server.CreateObject("G3MD");\n`;
                        js += `        __g3md_${comp.id}.Unsafe = ${comp.unsafe ? 'true' : 'false'};\n`;
                        js += `        ${comp.id}_val = __g3md_${comp.id}.Process(__g3files_${comp.id}.Read(__mdpath_${comp.id}, "utf-8"));\n`;
                        js += `    } else { ${comp.id}_val = ""; }\n`;
                    } else {
                        js += `    ${comp.id}_val = "";\n`;
                    }
                    js += `}\n`;
                } else {
                    const defaultVal = getStateDefaultValue(comp).replace(/"/g, '\\"');
                    js += `if (${comp.id}_val === null || ${comp.id}_val === "") { ${comp.id}_val = "${defaultVal}"; }\n`;
                }
            }
            return js;
        };

        const generateRenderHelpers = (compList) => {
            syncTableCells(compList);
            const renderables = collectRenderableComponents(compList, []);
            if (renderables.length === 0) return '';

            const generateAppendCodeForComponents = (compList, varName = '__html') => {
                let js = "";
                for (const comp of compList) {
                    if (comp.reRender && isRenderableAsyncType(comp.type)) {
                        const fn = getRenderFunctionName(comp.id);
                        js += `    ${varName} += ${fn}();\n`;
                        continue;
                    }
                    const { attrs, styleStr } = buildComponentAttrs(comp, { includeRuntimeBindings: true, skipClass: comp.type === 'modal', skipStyle: comp.type === 'modal' });
                    
                    if (comp.type === 'panel') {
                        js += `    ${varName} += '<div ${escapeJsSingleQuotedString(attrs)}>';\n`;
                        if (comp.children && comp.children.length > 0) {
                            js += generateAppendCodeForComponents(comp.children, varName);
                        }
                        js += `    ${varName} += '</div>';\n`;
                    } else if (comp.type === 'button') {
                        js += `    ${varName} += '<button ${escapeJsSingleQuotedString(attrs)}>${escapeJsSingleQuotedString(comp.text || '')}</button>';\n`;
                    } else if (comp.type === 'input') {
                        const inputType = escapeJsSingleQuotedString(comp.inputType || 'text');
                        const valExpr = comp.reRender ? `(${comp.id}_val !== null && ${comp.id}_val !== "" ? ${comp.id}_val : "${escapeJsSingleQuotedString(getStateDefaultValue(comp))}")` : `"${escapeJsSingleQuotedString(comp.text || '')}"`;
                        js += `    ${varName} += '<input type="${inputType}" ${escapeJsSingleQuotedString(attrs)} value="' + ${valExpr} + '">';\n`;
                    } else if (comp.type === 'textarea') {
                        const valExpr = comp.reRender ? `(${comp.id}_val !== null && ${comp.id}_val !== "" ? ${comp.id}_val : "${escapeJsSingleQuotedString(getStateDefaultValue(comp))}")` : `"${escapeJsSingleQuotedString(comp.text || '')}"`;
                        js += `    ${varName} += '<textarea ${escapeJsSingleQuotedString(attrs)}>' + ${valExpr} + '</textarea>';\n`;
                    } else if (comp.type === 'checkbox') {
                        const valExpr = comp.reRender ? `${comp.id}_val` : `${comp.checked}`;
                        js += `    var __checked = (${valExpr} === 'true' || ${valExpr} === true) ? ' checked' : '';\n`;
                        js += `    ${varName} += '<label ${escapeJsSingleQuotedString(attrs)}><input type="checkbox" value="${escapeJsSingleQuotedString(comp.value || '1')}"' + __checked + '> ${escapeJsSingleQuotedString(comp.text || '')}</label>';\n`;
                    } else if (comp.type === 'checkboxlist') {
                        const items = (comp.items || '').split(',').map(x => x.trim()).filter(Boolean);
                        const valExpr = comp.reRender ? `${comp.id}_val` : `""`;
                        js += `    var __selected = String(${valExpr} || '').split(',');\n`;
                        js += `    var __map = {};\n`;
                        js += `    for (var __i = 0; __i < __selected.length; __i++) { __map[__selected[__i].trim()] = true; }\n`;
                        js += `    ${varName} += '<div ${escapeJsSingleQuotedString(attrs)} data-g3al-type="checkboxlist">';\n`;
                        for (const itm of items) {
                            const lit = escapeJsSingleQuotedString(itm);
                            js += `    ${varName} += '<label><input type="checkbox" value="${lit}"' + (__map['${lit}'] ? ' checked' : '') + '> ${lit}</label><br>';\n`;
                        }
                        js += `    ${varName} += '</div>';\n`;
                    } else if (comp.type === 'radio') {
                        const valExpr = comp.reRender ? `${comp.id}_val` : `${comp.checked}`;
                        js += `    var __checked = (${valExpr} === 'true' || ${valExpr} === true) ? ' checked' : '';\n`;
                        js += `    ${varName} += '<label ${escapeJsSingleQuotedString(attrs)}><input type="radio" name="${escapeJsSingleQuotedString(comp.name || '')}" value="${escapeJsSingleQuotedString(comp.value || '1')}"' + __checked + '> ${escapeJsSingleQuotedString(comp.text || '')}</label>';\n`;
                    } else if (comp.type === 'radiobuttonlist') {
                        const items = (comp.items || '').split(',').map(x => x.trim()).filter(Boolean);
                        const name = escapeJsSingleQuotedString(comp.name || '');
                        const valExpr = comp.reRender ? `${comp.id}_val` : `""`;
                        js += `    ${varName} += '<div ${escapeJsSingleQuotedString(attrs)} data-g3al-type="radiobuttonlist">';\n`;
                        for (const itm of items) {
                            const lit = escapeJsSingleQuotedString(itm);
                            js += `    ${varName} += '<label><input type="radio" name="${name}" value="${lit}"' + (String(${valExpr} || '') === '${lit}' ? ' checked' : '') + '> ${lit}</label><br>';\n`;
                        }
                        js += `    ${varName} += '</div>';\n`;
                    } else if (comp.type === 'select') {
                        const opts = (comp.options || '').split(',').map(x => x.trim()).filter(Boolean);
                        const valExpr = comp.reRender ? `String(${comp.id}_val || "")` : `""`;
                        js += `    ${varName} += '<select ${escapeJsSingleQuotedString(attrs)}>';\n`;
                        for (const opt of opts) {
                            const lit = escapeJsSingleQuotedString(opt);
                            js += `    ${varName} += '<option value="${lit}"' + (${valExpr} === '${lit}' ? ' selected' : '') + '>${lit}</option>';\n`;
                        }
                        js += `    ${varName} += '</select>';\n`;
                    } else if (comp.type === 'listbox') {
                        const opts = (comp.options || '').split(',').map(x => x.trim()).filter(Boolean);
                        const multi = comp.multiSelect ? ' multiple' : '';
                        const size = comp.size ? ` size="${escapeJsSingleQuotedString(String(comp.size))}"` : '';
                        const valExpr = comp.reRender ? `${comp.id}_val` : `""`;
                        js += `    var __selected = String(${valExpr} || '').split(',');\n`;
                        js += `    var __map = {};\n`;
                        js += `    for (var __i = 0; __i < __selected.length; __i++) { __map[__selected[__i].trim()] = true; }\n`;
                        js += `    ${varName} += '<select ${escapeJsSingleQuotedString(attrs)}${multi}${size}>';\n`;
                        for (const opt of opts) {
                            const lit = escapeJsSingleQuotedString(opt);
                            if (comp.multiSelect) {
                                js += `    ${varName} += '<option value="${lit}"' + (__map['${lit}'] ? ' selected' : '') + '>${lit}</option>';\n`;
                            } else {
                                const valCheck = comp.reRender ? `String(${comp.id}_val || '')` : `""`;
                                js += `    ${varName} += '<option value="${lit}"' + (${valCheck} === '${lit}' ? ' selected' : '') + '>${lit}</option>';\n`;
                            }
                        }
                        js += `    ${varName} += '</select>';\n`;
                    } else if (comp.type === 'label') {
                        const valExpr = comp.reRender ? `(${comp.id}_val !== null && ${comp.id}_val !== "" ? ${comp.id}_val : "${escapeJsSingleQuotedString(getStateDefaultValue(comp))}")` : `"${escapeJsSingleQuotedString(comp.text || '')}"`;
                        js += `    ${varName} += '<span ${escapeJsSingleQuotedString(attrs)}>' + ${valExpr} + '</span>';\n`;
                    } else if (comp.type === 'hr') {
                        js += `    ${varName} += '<hr ${escapeJsSingleQuotedString(attrs)}>';\n`;
                    } else if (comp.type === 'fileuploader') {
                        const fileInputId = `${comp.id}_file`;
                        const resultId = `${comp.id}_result`;
                        const uploadAction = buildUploaderAction(comp.id, fileInputId);
                        const inputOnChange = comp.showUploadButton ? '' : ` onchange="${uploadAction}"`;
                        const uploadButton = comp.showUploadButton
                            ? `<button type="button" class="btn btn-primary" style="margin-top:5px;" onclick="${uploadAction}">${escapeHtmlText(comp.uploadButtonText || 'Send File')}</button>`
                            : '';
                        const valExpr = comp.reRender ? `(${comp.id}_val || "Ready")` : `"Ready"`;
                        js += `    ${varName} += '<div ${escapeJsSingleQuotedString(attrs)}><div class="sidebar-header" style="font-size:10px; margin-bottom:5px;">File Upload</div><input type="file" id="${escapeJsSingleQuotedString(fileInputId)}"${escapeJsSingleQuotedString(inputOnChange)}>${escapeJsSingleQuotedString(uploadButton)}<div id="${escapeJsSingleQuotedString(resultId)}" style="font-size:10px; color:#666; margin-top:5px;">Result: ' + ${valExpr} + '</div></div>';\n`;
                    } else if (comp.type === 'mdviewer') {
                        const valExpr = comp.reRender ? `(${comp.id}_val || '')` : `""`;
                        js += `    ${varName} += '<div ${escapeJsSingleQuotedString(attrs)}>' + ${valExpr} + '</div>';\n`;
                    } else if (comp.type === 'hiddenfield') {
                        const valExpr = comp.reRender ? `(${comp.id}_val !== null && ${comp.id}_val !== "" ? ${comp.id}_val : "${escapeJsSingleQuotedString(getStateDefaultValue(comp))}")` : `"${escapeJsSingleQuotedString(comp.value || '')}"`;
                        js += `    ${varName} += '<input type="hidden" ${escapeJsSingleQuotedString(attrs)} value="' + ${valExpr} + '">';\n`;
                    } else if (comp.type === 'image') {
                        js += `    ${varName} += '<img src="${escapeJsSingleQuotedString(comp.src || '')}" ${escapeJsSingleQuotedString(attrs)} alt="${escapeJsSingleQuotedString(comp.alt || '')}" title="${escapeJsSingleQuotedString(comp.title || '')}">';\n`;
                    } else if (comp.type === 'iframe') {
                        js += `    ${varName} += '<iframe src="${escapeJsSingleQuotedString(comp.src || '')}" ${escapeJsSingleQuotedString(attrs)} frameborder="0"></iframe>';\n`;
                    } else if (comp.type === 'link') {
                        js += `    ${varName} += '<a href="${escapeJsSingleQuotedString(comp.src || '#')}" ${escapeJsSingleQuotedString(attrs)}>${escapeJsSingleQuotedString(comp.text || '')}</a>';\n`;
                    } else if (comp.type === 'table') {
                        js += `    ${varName} += '<table ${escapeJsSingleQuotedString(attrs)}>';\n`;
                        if (comp.showHeader) {
                            js += `    ${varName} += '<thead><tr>';\n`;
                            for (let c = 1; c <= comp.cols; c++) {
                                let cell = comp.cells['h1_' + c];
                                const txt = (cell && cell.text) || 'Header ' + c;
                                js += `    ${varName} += '<th>${escapeJsSingleQuotedString(txt)}</th>';\n`;
                            }
                            js += `    ${varName} += '</tr></thead>';\n`;
                        }
                        js += `    ${varName} += '<tbody>';\n`;
                        for (let r = 1; r <= comp.rows; r++) {
                            js += `    ${varName} += '<tr>';\n`;
                            for (let c = 1; c <= comp.cols; c++) {
                                let cell = comp.cells[r + '_' + c];
                                js += `    ${varName} += '<td style="padding:5px;">';\n`;
                                if (cell) {
                                    if (cell.text) {
                                        js += `    ${varName} += '${escapeJsSingleQuotedString(cell.text)}';\n`;
                                    }
                                    if (cell.children && cell.children.length > 0) {
                                        js += generateAppendCodeForComponents(cell.children, varName);
                                    }
                                }
                                js += `    ${varName} += '</td>';\n`;
                            }
                            js += `    ${varName} += '</tr>';\n`;
                        }
                        js += `    ${varName} += '</tbody>';\n`;
                        if (comp.showFooter) {
                            js += `    ${varName} += '<tfoot><tr>';\n`;
                            for (let c = 1; c <= comp.cols; c++) {
                                let cell = comp.cells['f1_' + c];
                                const txt = (cell && cell.text) || 'Footer ' + c;
                                js += `    ${varName} += '<td>${escapeJsSingleQuotedString(txt)}</td>';\n`;
                            }
                            js += `    ${varName} += '</tr></tfoot>';\n`;
                        }
                        js += `    ${varName} += '</table>';\n`;
                    } else if (comp.type === 'placeholder') {
                        const valExpr = comp.reRender ? `(${comp.id}_val !== null && ${comp.id}_val !== "" ? ${comp.id}_val : "${escapeJsSingleQuotedString(getStateDefaultValue(comp))}")` : `"${escapeJsSingleQuotedString(comp.text || '')}"`;
                        js += `    ${varName} += '<div ${escapeJsSingleQuotedString(attrs)}>' + ${valExpr} + '</div>';\n`;
                    } else if (comp.type === 'rawhtml') {
                        js += `    ${varName} += '${escapeJsSingleQuotedString(comp.text || '')}';\n`;
                    }
                }
                return js;
            };

            let code = '// Shared render helpers used by both initial HTML and async re-render\n';

            for (const comp of renderables) {
                const fn = getRenderFunctionName(comp.id);
                const valExpr = `(${comp.id}_val !== null && ${comp.id}_val !== "" ? ${comp.id}_val : "${escapeJsSingleQuotedString(getStateDefaultValue(comp))}")`;
                const { attrs, styleStr } = buildComponentAttrs(comp, { includeRuntimeBindings: true, skipClass: comp.type === 'modal', skipStyle: comp.type === 'modal' });
                code += `function ${fn}() {\n`;

                if (comp.type === 'label') {
                    code += `    return '<span ${escapeJsSingleQuotedString(attrs)}>' + ${valExpr} + '</span>';\n`;
                } else if (comp.type === 'button') {
                    code += `    return '<button ${escapeJsSingleQuotedString(attrs)}>${escapeJsSingleQuotedString(comp.text || '')}</button>';\n`;
                } else if (comp.type === 'input') {
                    const inputType = escapeJsSingleQuotedString(comp.inputType || 'text');
                    code += `    return '<input type="${inputType}" ${escapeJsSingleQuotedString(attrs)} value="' + ${valExpr} + '">';\n`;
                } else if (comp.type === 'textarea') {
                    code += `    return '<textarea ${escapeJsSingleQuotedString(attrs)}>' + ${valExpr} + '</textarea>';\n`;
                } else if (comp.type === 'checkbox') {
                    code += `    var __checked = (${comp.id}_val === 'true' || ${comp.id}_val === true) ? ' checked' : '';\n`;
                    code += `    return '<label ${escapeJsSingleQuotedString(attrs)}><input type="checkbox" value="${escapeJsSingleQuotedString(comp.value || '1')}"' + __checked + '> ${escapeJsSingleQuotedString(comp.text || '')}</label>';\n`;
                } else if (comp.type === 'checkboxlist') {
                    const items = (comp.items || '').split(',').map(x => x.trim()).filter(Boolean);
                    code += `    var __selected = String(${comp.id}_val || '').split(',');\n`;
                    code += `    var __map = {};\n`;
                    code += `    for (var __i = 0; __i < __selected.length; __i++) { __map[__selected[__i].trim()] = true; }\n`;
                    code += `    var __html = '<div ${escapeJsSingleQuotedString(attrs)} data-g3al-type="checkboxlist">';\n`;
                    for (const itm of items) {
                        const lit = escapeJsSingleQuotedString(itm);
                        code += `    __html += '<label><input type="checkbox" value="${lit}"' + (__map['${lit}'] ? ' checked' : '') + '> ${lit}</label><br>';\n`;
                    }
                    code += `    __html += '</div>';\n`;
                    code += `    return __html;\n`;
                } else if (comp.type === 'radio') {
                    code += `    var __checked = (${comp.id}_val === 'true' || ${comp.id}_val === true) ? ' checked' : '';\n`;
                    code += `    return '<label ${escapeJsSingleQuotedString(attrs)}><input type="radio" name="${escapeJsSingleQuotedString(comp.name || '')}" value="${escapeJsSingleQuotedString(comp.value || '1')}"' + __checked + '> ${escapeJsSingleQuotedString(comp.text || '')}</label>';\n`;
                } else if (comp.type === 'radiobuttonlist') {
                    const items = (comp.items || '').split(',').map(x => x.trim()).filter(Boolean);
                    const name = escapeJsSingleQuotedString(comp.name || '');
                    code += `    var __html = '<div ${escapeJsSingleQuotedString(attrs)} data-g3al-type="radiobuttonlist">';\n`;
                    for (const itm of items) {
                        const lit = escapeJsSingleQuotedString(itm);
                        code += `    __html += '<label><input type="radio" name="${name}" value="${lit}"' + (String(${comp.id}_val || '') === '${lit}' ? ' checked' : '') + '> ${lit}</label><br>';\n`;
                    }
                    code += `    __html += '</div>';\n`;
                    code += `    return __html;\n`;
                } else if (comp.type === 'select') {
                    const opts = (comp.options || '').split(',').map(x => x.trim()).filter(Boolean);
                    code += `    var __html = '<select ${escapeJsSingleQuotedString(attrs)}>';\n`;
                    for (const opt of opts) {
                        const lit = escapeJsSingleQuotedString(opt);
                        code += `    __html += '<option value="${lit}"' + (String(${comp.id}_val || '') === '${lit}' ? ' selected' : '') + '>${lit}</option>';\n`;
                    }
                    code += `    __html += '</select>';\n`;
                    code += `    return __html;\n`;
                } else if (comp.type === 'listbox') {
                    const opts = (comp.options || '').split(',').map(x => x.trim()).filter(Boolean);
                    const multi = comp.multiSelect ? ' multiple' : '';
                    const size = comp.size ? ` size="${escapeJsSingleQuotedString(String(comp.size))}"` : '';
                    code += `    var __selected = String(${comp.id}_val || '').split(',');\n`;
                    code += `    var __map = {};\n`;
                    code += `    for (var __i = 0; __i < __selected.length; __i++) { __map[__selected[__i].trim()] = true; }\n`;
                    code += `    var __html = '<select ${escapeJsSingleQuotedString(attrs)}${multi}${size}>';\n`;
                    for (const opt of opts) {
                        const lit = escapeJsSingleQuotedString(opt);
                        if (comp.multiSelect) {
                            code += `    __html += '<option value="${lit}"' + (__map['${lit}'] ? ' selected' : '') + '>${lit}</option>';\n`;
                        } else {
                            code += `    __html += '<option value="${lit}"' + (String(${comp.id}_val || '') === '${lit}' ? ' selected' : '') + '>${lit}</option>';\n`;
                        }
                    }
                    code += `    __html += '</select>';\n`;
                    code += `    return __html;\n`;
                } else if (comp.type === 'hiddenfield') {
                    code += `    return '<input type="hidden" ${escapeJsSingleQuotedString(attrs)} value="' + ${valExpr} + '">';\n`;
                } else if (comp.type === 'fileuploader') {
                    const fileInputId = `${comp.id}_file`;
                    const resultId = `${comp.id}_result`;
                    const uploadAction = buildUploaderAction(comp.id, fileInputId);
                    const inputOnChange = comp.showUploadButton ? '' : ` onchange="${uploadAction}"`;
                    const uploadButton = comp.showUploadButton
                        ? `<button type="button" class="btn btn-primary" style="margin-top:5px;" onclick="${uploadAction}">${escapeHtmlText(comp.uploadButtonText || 'Send File')}</button>`
                        : '';
                    code += `    var __html = '<div ${escapeJsSingleQuotedString(attrs)}><div class="sidebar-header" style="font-size:10px; margin-bottom:5px;">File Upload</div><input type="file" id="${escapeJsSingleQuotedString(fileInputId)}"${escapeJsSingleQuotedString(inputOnChange)}>${escapeJsSingleQuotedString(uploadButton)}<div id="${escapeJsSingleQuotedString(resultId)}" style="font-size:10px; color:#666; margin-top:5px;">Result: ' + (${comp.id}_val || "Ready") + '</div></div>';\n`;
                    code += `    return __html;\n`;
                } else if (comp.type === 'mdviewer') {
                    code += `    return '<div ${escapeJsSingleQuotedString(attrs)}>' + (${comp.id}_val || '') + '</div>';\n`;
                } else if (comp.type === 'modal') {
                    const mClass = comp.cssClass ? ` ${comp.cssClass}` : '';
                    const winClass = escapeJsSingleQuotedString(`window${mClass}`);
                    const modalStyle = escapeJsSingleQuotedString(`display:none; ${styleStr}`);
                    code += `    var __html = '<div ${escapeJsSingleQuotedString(attrs)} class="${winClass}" style="${modalStyle}">';\n`;
                    code += `    __html += '<div class="window-header"><span>${escapeJsSingleQuotedString(comp.title || '')}</span><span style="cursor:pointer" onclick="G3AxonLive.closeModal(\\'${escapeJsSingleQuotedString(comp.id)}\\')">X</span></div>';\n`;
                    code += `    __html += '<div class="window-body">';\n`;
                    if (comp.modalType !== 'none') {
                        code += `    __html += '<div class="alert alert-${escapeJsSingleQuotedString(comp.modalType)}">' + ${valExpr} + '</div>';\n`;
                    } else {
                        code += `    __html += '<div>' + ${valExpr} + '</div>';\n`;
                    }
                    code += `    __html += '<div style="margin-top: 15px; display:flex; justify-content:flex-end; gap:5px;">';\n`;
                    if (comp.showBtn1) code += `    __html += '<button class="btn btn-primary" onclick="${escapeJsSingleQuotedString((comp.btn1Action || '').replace(/"/g, '&quot;'))}">${escapeJsSingleQuotedString(comp.btn1Text || '')}</button>';\n`;
                    if (comp.showBtn2) code += `    __html += '<button class="btn btn-secondary" onclick="${escapeJsSingleQuotedString((comp.btn2Action || '').replace(/"/g, '&quot;'))}">${escapeJsSingleQuotedString(comp.btn2Text || '')}</button>';\n`;
                    if (comp.showBtn3) code += `    __html += '<button class="btn btn-secondary" onclick="${escapeJsSingleQuotedString((comp.btn3Action || '').replace(/"/g, '&quot;'))}">${escapeJsSingleQuotedString(comp.btn3Text || '')}</button>';\n`;
                    code += `    __html += '</div></div></div>';\n`;
                    code += `    return __html;\n`;
                } else if (comp.type === 'placeholder') {
                    code += `    return '<div ${escapeJsSingleQuotedString(attrs)}>' + ${valExpr} + '</div>';\n`;
                } else if (comp.type === 'table') {
                    code += `    var __html = '<table ${escapeJsSingleQuotedString(attrs)}>';\n`;
                    if (comp.showHeader) {
                        code += `    __html += '<thead><tr>';\n`;
                        for (let c = 1; c <= comp.cols; c++) {
                            let cell = comp.cells['h1_' + c];
                            const txt = (cell && cell.text) || 'Header ' + c;
                            code += `    __html += '<th>${escapeJsSingleQuotedString(txt)}</th>';\n`;
                        }
                        code += `    __html += '</tr></thead>';\n`;
                    }
                    code += `    __html += '<tbody>';\n`;
                    for (let r = 1; r <= comp.rows; r++) {
                        code += `    __html += '<tr>';\n`;
                        for (let c = 1; c <= comp.cols; c++) {
                            let cell = comp.cells[r + '_' + c];
                            code += `    __html += '<td style="padding:5px;">';\n`;
                            if (cell) {
                                if (cell.text) {
                                    code += `    __html += '${escapeJsSingleQuotedString(cell.text)}';\n`;
                                }
                                if (cell.children && cell.children.length > 0) {
                                    code += generateAppendCodeForComponents(cell.children, '__html');
                                }
                            }
                            code += `    __html += '</td>';\n`;
                        }
                        code += `    __html += '</tr>';\n`;
                    }
                    code += `    __html += '</tbody>';\n`;
                    if (comp.showFooter) {
                        code += `    __html += '<tfoot><tr>';\n`;
                        for (let c = 1; c <= comp.cols; c++) {
                            let cell = comp.cells['f1_' + c];
                            const txt = (cell && cell.text) || 'Footer ' + c;
                            code += `    __html += '<td>${escapeJsSingleQuotedString(txt)}</td>';\n`;
                        }
                        code += `    __html += '</tr></tfoot>';\n`;
                    }
                    code += `    __html += '</table>';\n`;
                    code += `    return __html;\n`;
                } else {
                    code += `    return '';\n`;
                }

                code += `}\n`;
            }

            return code;
        };

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
                if (comp.reRender && isRenderableAsyncType(comp.type)) {
                    const fn = getRenderFunctionName(comp.id);
                    js += `    AxonLive.RegisterComponent("${comp.id}", ${fn}());\n`;
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

        const generateEventSwitch = (compList) => {
            let js = "";
            for (const comp of compList) {
                if (comp.events && Object.keys(comp.events).length > 0 || comp.type === 'fileuploader') {
                    js += `        case "${comp.id}":\n`;
                    js += `            // Granular API: Auto-generated proxies for interaction\n`;
                    js += `            var ${comp.id} = AxonLive.GetComponent("${comp.id}");\n`;

                    if (comp.type === 'fileuploader') {
                        js += `            // Multipart File Upload Logic\n`;
                        js += `            if (evtName === "onupload" && Request.TotalBytes > 0 && Request.ServerVariables("HTTP_X_G3AL_UPLOAD") == "true") {\n`;
                        js += `                var uploader = Server.CreateObject("G3FILEUPLOADER");\n`;
                        js += `                uploader.MaxFileSize = ${comp.maxFileSize || 5242880};\n`;
                        if (comp.allowedExtensions) {
                            (comp.allowedExtensions.split(',')).forEach(ext => {
                                js += `                uploader.AllowExtension("${ext.trim()}");\n`;
                            });
                        }
                        if (comp.blockedExtensions) {
                            (comp.blockedExtensions.split(',')).forEach(ext => {
                                js += `                uploader.BlockExtension("${ext.trim()}");\n`;
                            });
                        }
                        const savedName = comp.savedFileName ? `"${comp.savedFileName}"` : '""';
                        const preserve = comp.savedFileName ? (comp.preserveName ? "true" : "false") : "true";
                        js += `                uploader.PreserveOriginalName = ${preserve};\n`;
                        js += `                var result = uploader.Process("file", "${comp.targetDir || '/uploads'}", ${savedName});\n`;
                        js += `                if (result.Item("IsSuccess")) { ${comp.id}_val = "Success: " + result.Item("RelativePath"); }\n`;
                        js += `                else { ${comp.id}_val = "Error: " + result.Item("ErrorMessage"); }\n`;
                        js += `            }\n`;
                    }

                    // Automatically sync _val from incoming event arguments if it's a stateful component
                    if (comp.reRender) {
                        js += `            if (AxonLive.GetEventArg("value") !== "") { ${comp.id}_val = AxonLive.GetEventArg("value"); }\n`;
                        if (comp.type === 'checkbox') {
                            js += `            ${comp.id}_val = AxonLive.GetEventArg("checked");\n`;
                        }
                    }

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
            const timers = allComponents.value.filter(c => c.type === 'timer');
            const scripts = allComponents.value.filter(c => c.type === 'script');
            const styles = allComponents.value.filter(c => c.type === 'style');

            let timerInitCode = "";
            for (const t of timers) {
                timerInitCode += `    // Initialize timer: ${t.id}\n`;
                timerInitCode += `    AxonLive.SetTimer("${t.id}", "${t.triggerEvent}", ${t.delay});\n`;
            }

            let styleBlock = "";
            for (const s of styles) styleBlock += `<style>\n${s.text}\n</style>\n`;

            let scriptBlock = "";
            for (const sc of scripts) {
                if (sc.serverSide) {
                    scriptBlock += `<script language="javascript" runat="server">\n${sc.text}\n</script>\n`;
                } else {
                    scriptBlock += `<script>\n${sc.text}\n</script>\n`;
                }
            }

            const switchLogic = generateEventSwitch(allComponents.value);
            const renderLogic = generateReRenderCalls(allComponents.value);
            const stateRestore = generateStateRestore(allComponents.value);
            const statePersist = generateStatePersist(allComponents.value);
            const renderHelpers = generateRenderHelpers(allComponents.value);
            const htmlLayout = generateHTML(components.value, "    ");
            const hasModal = hasComponentType(allComponents.value, 'modal');

            // Generate hidden field HTML separately (not on canvas but emitted into the page body)
            const hiddenFields = allComponents.value.filter(c => c.type === 'hiddenfield');
            let hiddenFieldsHtml = "";
            for (const hf of hiddenFields) {
                const { attrs: hfAttrs } = buildComponentAttrs(hf, { includeRuntimeBindings: true });
                hiddenFieldsHtml += `    <input type="hidden" ${hfAttrs} value="${escapeHtmlAttr(hf.value || '')}">${'\n'}`;
            }

            let mainContainerStyle = ``;
            if (pageSettings.display === 'flex') {
                mainContainerStyle = ` style="display:flex; flex-direction:${pageSettings.flexDirection}; justify-content:${pageSettings.justifyContent}; align-items:${pageSettings.alignItems}; width:100%; height:100%;"`;
            }

            return `<%@ Language="JavaScript" %>
<%
/* Auto-generated by G3pix AxonLive Visual Builder */

${hasModal ? `/*
 * MODAL MANAGEMENT TIPS:
 * - To show a modal:    G3AxonLive.showModal("modalID");
 * - To close a modal:   G3AxonLive.closeModal("modalID");
 * - To toggle a modal:  G3AxonLive.toggleModal("modalID");
 * 
 * You can call these from any client-side onclick handler or via 
 * AxonLive.Trigger("btnID", "onclick") from the server.
 */
` : ''}

var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

var sessionID = Session.SessionID;

${stateRestore}
${renderHelpers}
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
${styleBlock}
</head>
<body>

<div id="main-container">
    <div id="content"${mainContainerStyle}>
${htmlLayout}${hiddenFieldsHtml}    </div>
</div>

${scriptBlock}

<script src="/axonlive/g3axonlive.js"></script>
<script>
    G3AxonLive.init();
</script>
</body>
</html>`;
        });

        watch(generatedCode, () => {
            highlightGeneratedCode();
        });

        const jsonTree = computed(() => {
            return JSON.stringify(allComponents.value, null, 4);
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
            fileInput,
            availableComponents,
            allComponents,
            components,
            functionalComponents,
            selectedComponent,
            pageSettings,
            showJsonTree,
            codePaneState,
            generatedCode,
            jsonTree,
            editableJsonTree,
            aspCodeEl,
            newEventName,
            newClientEventName,
            updateFromJson,
            toggleCodePane,
            onDragStart,
            onDrop,
            selectComponent,
            removeComponent,
            moveComponent,
            addEvent,
            addClientEvent,
            clearCanvas,
            copyCode,
            downloadCode,
            handleDragStart,
            handleResizeStart,
            historyIndex,
            history,
            undo,
            redo,
            copyComp,
            pasteComp,
            duplicateComp,
            clipboard,
            contextMenu,
            onContextMenu,
            saveModel,
            openModel,
            handleFileOpen
        };
    }
});

app.mount('#app');