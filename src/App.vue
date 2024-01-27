<template>
	<div class="card">
    <Splitter  class="mb-5">
      <SplitterPanel class="flex flex-column" :size="25">
        <Toolbar class="p-mb-4 excel-toolbar">
          <template #start>
            <Button label="Новый" 
              icon="pi pi-plus" 
              class="p-button-success p-mr-2" 
              @click="openNew" />
          </template>
        </Toolbar>
        <Menu :model="items" class="excel-calc-list"/>
        
      </SplitterPanel>
      <SplitterPanel class="flex flex-column" :size="75"> 
        <ExcelWorkbook v-model="EWBData" v-show="ExcelWorkbookvisible"/>
      </SplitterPanel>
    </Splitter>
    <Toast />

    <Dialog v-model:visible="newWorkBookDialog" 
      :style="{width: '450px'}" 
      header="Новая книга эксель" 
      :modal="true" 
      class="p-fluid" >
      
      <div class="p-field">
          <label for="name">Имя</label>
          <InputText id="name" v-model.trim="WorkBook.name" required="true" :class="{'p-invalid': submitted && !WorkBook.name}" />
          <small class="p-error" v-if="submitted && !WorkBook.name">Имя требуется.</small>
      </div>

      <template #footer>
          <Button label="Отмена" icon="pi pi-times" class="p-button-text" @click="hideDialog"/>
          <Button label="Сохранить" icon="pi pi-check" class="p-button-text" @click="saveWorkBook" />
      </template>
    </Dialog>
  </div>
  
</template>

<script setup>
  import { ref, onMounted } from 'vue';
  import axios from 'axios'
  import Splitter from 'primevue/splitter';
  import SplitterPanel from 'primevue/splitterpanel';
  import Menu from 'primevue/menu';
  import { useToast } from "primevue/usetoast";
  import Toolbar from 'primevue/toolbar';
  import Button from 'primevue/button';
  import Dialog from 'primevue/dialog';
  import InputText from 'primevue/inputtext'
  import ExcelWorkbook from './components/ExcelWorkbook.vue'

  const toast = useToast();
  const WorkBook = ref({});
  const EWBData = ref({});
  const items = ref([]);
  const point = '/api/excelcalc_workbooks'
  const newWorkBookDialog = ref(false);
  const submitted = ref(false);
  const ExcelWorkbookvisible = ref(false);
  onMounted(() => {
    //loading.value = true;
    loadLazyData();
  });
  const loadLazyData = (event) => {
    let params = { 
      limit: 0, 
      setTotal: 0, 
    }
    axios.get(point,{ params: params})
    .then(function (response) {
      // console.log(response.data);
      items.value = [];
      response.data.rows.forEach((workbook) => {
        workbook.label = workbook.name;
        workbook.command = (e) => {
          // console.log(e.item.id);
          EWBData.value = e.item;
          ExcelWorkbookvisible.value = true;
          // toast.add({ severity: 'success', summary: 'Success', detail: 'File created', life: 3000 });
        }
        items.value.push(workbook);
        
      });
      EWBData.value = items.value[0];
      ExcelWorkbookvisible.value = true;
    });
  };
  
  const openNew = () => 
  {
    WorkBook.value = {};
    submitted.value = false;
    newWorkBookDialog.value = true;
  };
  const hideDialog = () => 
  {
    newWorkBookDialog.value = false;
  };
  const saveWorkBook = () => 
  {
    submitted.value = true;
    if(!WorkBook.value.name) return; 
    axios.put(point, WorkBook.value)
      .then((response) =>{ 
          loadLazyData();
      });

    newWorkBookDialog.value = false;
    WorkBook.value = {};
  };
</script>
