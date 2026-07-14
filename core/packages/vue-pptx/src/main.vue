<script>
import { defineComponent, ref, onMounted, onBeforeUnmount, nextTick, watch } from 'vue-demi';
import {init as initPptxPreviewer} from 'pptx-preview';

export default defineComponent({
    name: 'VueOfficePptx',
    props: {
        src: [String, ArrayBuffer, Blob],
        requestOptions: {
            type: Object,
            default: () => ({})
        },
        options: {
            type: Object,
            default: () => ({})
        }
    },
    emits: ['rendered', 'error'],
    setup(props, { emit }) {
        let pptxViewer = null;
        let renderVersion = 0;
        let destroyed = false;
        let previewQueue = Promise.resolve();
        let requestController = null;
        const rootRef = ref(null);

        function cancelRequest(){
            if (requestController) {
                requestController.abort();
                requestController = null;
            }
        }

        function init(){
            if (pptxViewer || !rootRef.value) {
                return;
            }
            let container = rootRef.value;
            let rect = container.getBoundingClientRect();
            let width = props.options.width || rect.width || 960;
            let height = props.options.height || rect.height || 540;
            pptxViewer = initPptxPreviewer(container, {
                width,
                height
            });
        }

        function getPptxData(src, requestOptions){
            if(typeof src === 'string'){
                return fetch(src, requestOptions).then(response=>{
                    if (!response.ok) {
                        throw new Error(`PPTX请求失败: ${response.status}`);
                    }
                    return response.arrayBuffer();
                });
            }
            if(src instanceof ArrayBuffer){
                return Promise.resolve(src);
            }
            if(src instanceof Blob){
                return src.arrayBuffer();
            }
            return Promise.reject(new Error('PPTX预览仅支持URL、ArrayBuffer或Blob'));
        }

        function clearPreview(){
            if (rootRef.value) {
                rootRef.value.innerHTML = '';
            }
        }

        function isCurrent(version){
            return !destroyed && version === renderVersion;
        }

        function preview(){
            const version = ++renderVersion;
            const src = props.src;
            cancelRequest();
            const controller = typeof AbortController === 'function' ? new AbortController() : null;
            requestController = controller;
            const requestOptions = controller
                ? {...props.requestOptions, signal: controller.signal}
                : props.requestOptions;
            clearPreview();
            if (!src) {
                requestController = null;
                return;
            }

            previewQueue = previewQueue.catch(() => {}).then(async () => {
                if (!isCurrent(version)) {
                    if (requestController === controller) {
                        requestController = null;
                    }
                    return;
                }
                const arrayBuffer = await getPptxData(src, requestOptions);
                if (requestController === controller) {
                    requestController = null;
                }
                if (!isCurrent(version)) {
                    return;
                }
                clearPreview();
                init();
                const pptx = await pptxViewer.preview(arrayBuffer);
                if (isCurrent(version)) {
                    emit('rendered', pptx);
                }
            }).catch(e=>{
                if (requestController === controller) {
                    requestController = null;
                }
                if (isCurrent(version)) {
                    clearPreview();
                    emit('error', e);
                }
            });
        }

        onMounted(() => {
            nextTick(() => {
                if (!destroyed) {
                    preview();
                }
            });
        });

        watch(() => props.src, () => {
            preview();
        });

        onBeforeUnmount(() => {
            destroyed = true;
            renderVersion += 1;
            cancelRequest();
            if (pptxViewer && typeof pptxViewer.destroy === 'function') {
                pptxViewer.destroy();
            }
            clearPreview();
            pptxViewer = null;
        });

        return {
            rootRef
        };
    }
});
</script>

<template>
    <div class="vue-office-pptx">
        <div class="vue-office-pptx-main" ref="rootRef" style="width:100%; height: 100%;"></div>
    </div>
</template>
<style lang="less">
</style>
