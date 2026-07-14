<script>
import {defineComponent, ref, onMounted, onBeforeUnmount, watch, nextTick} from 'vue-demi';
import Spreadsheet from './x-spreadsheet/index';
import {getData, readExcelData, transferExcelToSpreadSheet} from './excel';
import {ExcelWorkerClient, getWorkerOptions} from './worker-client';
import {renderImage, clearCache} from './media';
import {debounce} from 'lodash';
import {download as downloadFile} from '../../../utils/url';

const defaultOptions = {
    xls: false,
    minColLength: 20
};
export default defineComponent({
    name: 'VueOfficeExcel',
    props: {
        src: [String, ArrayBuffer, Blob],
        requestOptions: {
            type: Object,
            default: () => ({})
        },
        options: {
            type: Object,
            default: () => ({
               ...defaultOptions
            })
        }
    },
    emits: ['rendered', 'error', 'switchSheet', 'cellSelected', 'cellsSelected'],
    setup(props, {emit}) {
        const wrapperRef = ref(null);
        const rootRef = ref(null);
        let workbookDataSource = {
            _worksheets:[]
        };
        let mediasSource = [];
        let sheetIndex = 0;
        let ctx = null;
        let xs = null;
        let offset = null;
        let fileData = null;
        let renderVersion = 0;
        let destroyed = false;
        let requestController = null;
        let workerClient = null;

        function getWorkerClient(){
            if (workerClient || typeof Worker !== 'function') {
                return workerClient;
            }
            try {
                workerClient = new ExcelWorkerClient(() => new Worker(
                    new URL('./excel.worker.js', import.meta.url),
                    {type: 'module'}
                ));
            } catch (e) {
                workerClient = null;
            }
            return workerClient;
        }

        function cancelRequest(){
            if (requestController) {
                requestController.abort();
                requestController = null;
            }
            workerClient && workerClient.cancel();
        }

        function isCurrent(version) {
            return !destroyed && version === renderVersion && xs;
        }

        function renderExcel(buffer, version) {
            if (!isCurrent(version)) {
                return;
            }
            fileData = buffer;
            const options = {...defaultOptions, ...props.options};
            const client = !props.options.beforeTransformData && getWorkerClient();
            const parsePromise = client
                ? client.parse(buffer, getWorkerOptions(options))
                : readExcelData(buffer, props.options.xls).then(workbook => ({workbook}));
            parsePromise.then(parsed => {
                if (!isCurrent(version)) {
                    return;
                }
                let workbookData;
                let medias;
                let workbookSource;
                if (parsed.workbook) {
                    let workbook = parsed.workbook;
                    if(props.options.beforeTransformData && typeof props.options.beforeTransformData === 'function' ){
                        workbook = props.options.beforeTransformData(workbook);
                    }
                    const result = transferExcelToSpreadSheet(workbook, options);
                    workbookData = result.workbookData;
                    medias = result.medias;
                    workbookSource = result.workbookSource;
                } else {
                    fileData = parsed.buffer || buffer;
                    workbookData = parsed.workbookData;
                    medias = parsed.medias;
                    workbookSource = parsed.workbookSource;
                }
                if(props.options.transformData && typeof props.options.transformData === 'function' ){
                    workbookData = props.options.transformData(workbookData);
                }
                mediasSource = medias;
                workbookDataSource = workbookSource;
                offset = null;
                sheetIndex = 0;
                clearCache();
                xs.loadData(workbookData);
                renderImage(ctx, mediasSource, workbookDataSource._worksheets[sheetIndex], offset, props.options);
                emit('rendered');
                emit('switchSheet', 0);
                //涉及clear和offset

            }).catch(e => {
                if (!isCurrent(version)) {
                    return;
                }
                console.warn(e);
                mediasSource = [];
                workbookDataSource = {
                    _worksheets:[]
                };
                clearCache();
                xs && xs.loadData({});
                emit('error', e);
                emit('switchSheet', 0);
            });
        }

        function loadSource(src) {
            const version = ++renderVersion;
            cancelRequest();
            if (!src) {
                return;
            }
            const controller = typeof AbortController === 'function' ? new AbortController() : null;
            requestController = controller;
            const requestOptions = controller
                ? {...props.requestOptions, signal: controller.signal}
                : props.requestOptions;
            getData(src, requestOptions).then(buffer => {
                if (requestController === controller) {
                    requestController = null;
                }
                renderExcel(buffer, version);
            }).catch(e => {
                if (requestController === controller) {
                    requestController = null;
                }
                if (!isCurrent(version)) {
                    return;
                }
                mediasSource = [];
                workbookDataSource = {
                    _worksheets:[]
                };
                xs.loadData({});
                emit('error', e);
            });
        }
        
        onMounted(() => {
            nextTick(()=>{
                xs = new Spreadsheet(rootRef.value, {
                    mode: 'read',
                    showToolbar: false,
                    showContextmenu: props.options.showContextmenu || false,
                    view: {
                        height: () => wrapperRef.value && wrapperRef.value.clientHeight || 300,
                        width: () => wrapperRef.value && wrapperRef.value.clientWidth || 1200,
                    },
                    row: {
                        height: 24,
                        len: 100
                    },
                    col: {
                        len: 26,
                        width: 80,
                        indexWidth: 60,
                        minWidth: 60,
                    },
                    autoFocus: false
                }).loadData({});

                xs.on('cell-selected', (cell, ri, ci) => {
                    emit('cellSelected', {
                        cell,
                        rowIndex: ri,
                        columnIndex: ci
                    });
                });
                xs.on('cells-selected', (cell, { sri, sci, eri, eci }) => {
                    emit('cellsSelected', {
                        cell,
                        startRowIndex: sri,
                        startColumnIndex: sci,
                        endRowIndex: eri,
                        endColumnIndex: eci
                    });
                });

                let swapFunc = xs.bottombar.swapFunc;
                xs.bottombar.swapFunc = function (index) {
                    swapFunc.call(xs.bottombar, index);
                    sheetIndex = index;
                    offset = xs.sheet.data.getSelectedRect();
                    setTimeout(()=>{
                        xs.reRender();
                        renderImage(ctx, mediasSource, workbookDataSource._worksheets[sheetIndex], offset, props.options);
                        emit('switchSheet', index);
                    });

                };

                let renderImageDebounce = debounce(renderImage, 200, {
                    leading: true
                });
                let tableRender = xs.sheet.table.render;
                xs.sheet.table.render = function (...args){
                    xs && xs.sheet && tableRender.apply(xs.sheet.table, args);
                    renderImageDebounce(ctx, mediasSource, workbookDataSource._worksheets[sheetIndex], offset, props.options);
                };

                // let clear = xs.sheet.editor.clear;
                // xs.sheet.editor.clear = function (...args){
                //     clear.apply(xs.sheet.editor, args);
                //     setTimeout(()=>{
                //         renderImage(ctx, mediasSource, workbookDataSource._worksheets[sheetIndex], offset, props.options);
                //     });
                // };
                let setOffset = xs.sheet.editor.setOffset;
                xs.sheet.editor.setOffset = function (...args){
                    setOffset.apply(xs.sheet.editor, args);
                    offset = args[0];
                };
                const canvas = rootRef.value.querySelector('canvas');
                ctx = canvas.getContext('2d');
                if (props.src) {
                    loadSource(props.src);
                }
            });
        });

        onBeforeUnmount(()=>{
            destroyed = true;
            renderVersion += 1;
            cancelRequest();
            workerClient && workerClient.destroy();
            workerClient = null;
            xs && xs.destroy();
            xs = null;
        });
        watch(() => props.src, () => {
            if (props.src) {
                loadSource(props.src);
            } else {
                renderVersion += 1;
                cancelRequest();
                mediasSource = [];
                workbookDataSource = {
                    _worksheets:[]
                };
                xs && xs.loadData({});
                emit('error', new Error('src属性不能为空'));
            }
        });
        function save(fileName){
            downloadFile(fileName || `vue-office-excel-${new Date().getTime()}.xlsx`,fileData);
        }
        return {
            wrapperRef,
            rootRef,
            save
        };
    }
});
</script>

<template>
    <div class="vue-office-excel" ref="wrapperRef">
        <div class="vue-office-excel-main" ref="rootRef"></div>
    </div>
</template>
<style lang="less">

</style>
