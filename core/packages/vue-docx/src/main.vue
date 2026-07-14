<script>
import {defineComponent, ref, onMounted, onBeforeUnmount, watch} from 'vue-demi';
import docx from './docx';
import {download as downloadFile} from '../../../utils/url';

export default defineComponent({
    name: 'VueOfficeDocx',
    props: {
        src: [String, ArrayBuffer, Blob],
        requestOptions: {
            type: Object,
            default: () => ({})
        },
        options:{
            type: Object,
            default: () => ({})
        }
    },
    emits: ['rendered', 'error'],
    setup(props, {emit}) {
        const rootRef = ref(null);
        let fileData = null;
        let renderVersion = 0;
        let destroyed = false;
        let requestController = null;

        function cancelRequest(){
            if (requestController) {
                requestController.abort();
                requestController = null;
            }
        }

        function isCurrent(version){
            return !destroyed && version === renderVersion;
        }

        function clearPreview(){
            if (rootRef.value) {
                rootRef.value.innerHTML = '';
            }
        }

        function init() {
            let container = rootRef.value;
            const version = ++renderVersion;
            const src = props.src;
            cancelRequest();
            if (!src || !container) {
                clearPreview();
                return;
            }
            const controller = typeof AbortController === 'function' ? new AbortController() : null;
            requestController = controller;
            const requestOptions = controller
                ? {...props.requestOptions, signal: controller.signal}
                : props.requestOptions;
            clearPreview();
            docx.getData(src, requestOptions).then(async res => {
                if (requestController === controller) {
                    requestController = null;
                }
                if (!isCurrent(version)) {
                    return;
                }
                fileData = await docx.getBlob(res);
                if (!isCurrent(version)) {
                    return;
                }
                return docx.render(fileData, container, props.options).then(() => {
                    if (!isCurrent(version)) {
                        return;
                    }
                    emit('rendered');
                }).catch(e => {
                    if (!isCurrent(version)) {
                        return;
                    }
                    clearPreview();
                    emit('error', e);
                });
            }).catch(e => {
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
            if (props.src) {
                init();
            }
        });

        watch(() => props.src, () => {
            if (props.src) {
                init();
            } else {
                renderVersion += 1;
                cancelRequest();
                clearPreview();
                emit('rendered');
            }
        });

        onBeforeUnmount(() => {
            destroyed = true;
            renderVersion += 1;
            cancelRequest();
            clearPreview();
        });
        function save(fileName){
            downloadFile(fileName || `vue-office-docx-${new Date().getTime()}.docx`,fileData);
        }
        return {
            rootRef,
            save
        };
    }
});
</script>

<template>
    <div class="vue-office-docx">
        <div class="vue-office-docx-main" ref="rootRef"></div>
    </div>
</template>

<style lang="less">
.vue-office-docx {
    height: 100%;
    overflow-y: auto;
    .docx-wrapper {
        > section.docx {
           margin-bottom: 5px;
        }
    }
}

@media screen and (max-width: 800px) {
    .vue-office-docx {
        .docx-wrapper {
            padding: 10px;

            > section.docx {
                padding: 10px !important;
                width: 100% !important;
            }
        }
    }
}

</style>
