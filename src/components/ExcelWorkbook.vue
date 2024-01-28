
<template>
	<div class="card p-fluid">
    <Toolbar class="p-mb-4 excel-toolbar">
      <template #start>
        <Button type="button" icon="pi pi-plus" @click="toggle" aria-haspopup="true" aria-controls="overlay_menu" />
        <Menu ref="menu" id="overlay_menu" :model="items" :popup="true" />
      </template>
      <template #end>
        <Button type="button" icon="pi pi-check" @click="save" label="Сохранить"/>
      </template>
    </Toolbar>
    <DataTable :value="lineItems" lazy  ref="dt" dataKey="A"
      :loading="loading" 
       tableStyle="min-width: 75rem"
      editMode="cell" @cell-edit-complete="onCellEditComplete"
      showGridlines
      :pt="{
        table: { style: 'min-width: 50rem' },
        column: {
            bodycell: ({ state }) => ({
                class: [{ 'pt-0 pb-0': state['d_editing'] }]
            })
        }
      }"
      >
      <Column field="id" header="id" style="width: 1px;padding: 1rem 1px 1rem 10px;">
        <template #body="{ data, field }">
            {{ data[field].v }} 
        </template>
      </Column>
      <Column v-for="col of columns" :key="col.field" :field="col.field" :header="col.header" style="width: 3rem;position: relative;">
        <template #body="{ data, field }">
            {{ data[field].v }} 
              <Button v-if="data[field].f" type="button" icon="pi pi-check" size="small" 
              class="excel-calc-check"
              style="position:absolute;top: 5px; right:0;" :data-calc="data[field].v"/>
        </template>
        <template #editor="{ data, field }">
          <template v-if="field != 'id'">
            <InputText v-if="data[field].f" v-model="data[field].f" onfocus="this.select()"/>
            <InputText v-else v-model="data[field].v" onfocus="this.select()"/>
          </template>
          <template v-else>
            {{ data[field].v }}
          </template>
        </template>
      </Column>
    </DataTable>
    
	</div>
</template>
<style>
  .excel-calc-check{
    display:none;
  }
</style>
<script setup>
  import { ref, computed, watch, onMounted } from 'vue';
  import DataTable from 'primevue/datatable'
  import Column from 'primevue/column'
  import InputText from 'primevue/inputtext'
  import Button from 'primevue/button';
  import Menu from 'primevue/menu';
  import Toolbar from 'primevue/toolbar';
  import Dialog from 'primevue/dialog';
  import axios from 'axios'
  import XLSX_CALC from 'xlsx-calc';
  import { useToast } from "primevue/usetoast";

  const toast = useToast();
  const props = defineProps({
    modelValue: {
      type: Object,
    }
  });
  const lineItems = ref([]);
  
  const excelColToInt = function (colName) {
    var digits = colName.toUpperCase().split(''),
      number = 0;

    for (var i = 0; i < digits.length; i++) {
      number += (digits[i].charCodeAt(0) - 64)*Math.pow(26, digits.length - i - 1);
    }

    return number;    
  };
  const intToExcelCol = function (number) {
    var colName = '',
      dividend = Math.floor(Math.abs(number)),
      rest;

    while (dividend > 0) {
      rest = (dividend - 1) % 26;
      colName = String.fromCharCode(65 + rest) + colName;
      dividend = parseInt((dividend - rest)/26);
    }
    return colName;
  };
  const columns = ref([
    {field: 'id', header: 'id'},

  ]);
  let workbook = ref({});
  const genlineItems = function (Sheet) {
    let maxCol = 0; let maxRow = 0;
    let rown; let col; let colName;
    for(let key in Sheet){
      rown = parseInt(key.match(/\d+/g)[0]);
      if(rown > maxRow) maxRow = rown;
      colName = key.replace(/[^a-zA-Z]+/g, '');
      col = excelColToInt(colName);
      if(col > maxCol) maxCol = col;
    }
    if(maxCol < 5) maxCol = 5;
    let row = {};
    columns.value = [];
    for(col = 1;col <= maxCol;col++){
      colName = intToExcelCol(col);
      columns.value.push({field: colName, header: colName});
    }
    lineItems.value = [];
    for(rown = 1;rown <= maxRow;rown++){
      row = {};
      row.id = {v:rown};
      for(col = 1;col <= maxCol;col++){
        colName = intToExcelCol(col);
        row[colName] = {v:""};
        if(Sheet.hasOwnProperty(colName + rown)){
          row[colName].v = Sheet[colName + rown].v;
          if(Sheet[colName + rown].hasOwnProperty('f')){
            row[colName].f = '=' + Sheet[colName + rown].f;
          }
        }
      }
      lineItems.value.push(row);
      loading.value = false;
    }
  };
  watch(() => props.modelValue, (newValue) => {
    // console.log('props.modelValue', newValue);
    if(!props.modelValue.workbook){
      props.modelValue.workbook = {Sheets: {Sheet1: {A1:{v:""}}}};
    }
    if(!props.modelValue.workbook.Sheets){
      props.modelValue.workbook = {Sheets: {Sheet1: {A1:{v:""}}}};
    }
    workbook = JSON.parse(JSON.stringify(props.modelValue.workbook));
    genlineItems(workbook.Sheets.Sheet1);
  });

  const point = '/api/excelcalc_workbooks'
  const dt = ref();
  const loading = ref(true);

  
  
  
  const onCellEditComplete = (event) => {
    let { data, newValue, field } = event;
    // console.log('onCellEditComplete',newValue,data[field].f);
    if(field == 'id') return;
    //props.modelValue.workbook.Sheets.Sheet1
    if(newValue.v != '' || newValue.f != ''){
      if(newValue.v[0] == '='){
        newValue.f = '=' + newValue.v.slice(1);
        newValue.v = '';
      }
      if(newValue.hasOwnProperty('f') && newValue.f[0] != '='){
        newValue.v = newValue.f;
        newValue.f = '';
      }
      if(!workbook.Sheets.Sheet1.hasOwnProperty(field + data.id.v)){
        workbook.Sheets.Sheet1[field + data.id.v] = {v:''};
      }
      if(newValue.hasOwnProperty('f')){
        if(newValue.f[0] == '='){
          workbook.Sheets.Sheet1[field + data.id.v].f = newValue.f.slice(1);
        }else{
          delete workbook.Sheets.Sheet1[field + data.id.v].f;
          workbook.Sheets.Sheet1[field + data.id.v].v = newValue.v;
        }
      }else if(workbook.Sheets.Sheet1[field + data.id.v].hasOwnProperty('f')){
        delete workbook.Sheets.Sheet1[field + data.id.v].f;
        workbook.Sheets.Sheet1[field + data.id.v].v = newValue.v;
      }else{
        workbook.Sheets.Sheet1[field + data.id.v].v = newValue.v;
      }
      // console.log('onCellEditComplete2',newValue)
      XLSX_CALC(workbook);
      genlineItems(workbook.Sheets.Sheet1);
    }
  };
  const loadLazyData = (event) => {
    loading.value = true;
    // lazyParams.value = { ...lazyParams.value, first: event?.first || first.value };
    // // console.log('lazyParams.value',lazyParams.value)
    // // console.log('event',event)
    // let params = { 
    //   limit: 10, 
    //   setTotal: 1, 
    //   offset:lazyParams.value.first,
    //   sortField:lazyParams.value.sortField,
    //   sortOrder:lazyParams.value.sortOrder,
    // }
    axios.get(point,{ params: params})
    .then(function (response) {
      // console.log(response.data);
      // console.log(response.status);
      // console.log(response.statusText);
      // console.log(response.headers);
      // console.log(response.config);
      // lineItems.value = response.data.rows;
      // totalRecords.value = response.data.total;
      loading.value = false;
    });
  };

  const menu = ref();
  const items = ref([
    {
      label: 'Строка',
      icon: 'pi pi-plus',
      command: () => {
        // toast.add({ severity: 'success', summary: 'Success', detail: 'File created', life: 3000 });
        let maxRow = lineItems.value.length + 1;
        workbook.Sheets.Sheet1['A'+ maxRow] = {v:''};
        genlineItems(workbook.Sheets.Sheet1);
      }
    },
    {
      label: 'Строка',
      icon: 'pi pi-minus',
      command: () => {
        // toast.add({ severity: 'success', summary: 'Success', detail: 'File created', life: 3000 });
        let maxRow = lineItems.value.length;
        let rown;
        for(let key in workbook.Sheets.Sheet1){
          rown = parseInt(key.match(/\d+/g)[0]);
          if(rown == maxRow) delete workbook.Sheets.Sheet1[key];
        }
        genlineItems(workbook.Sheets.Sheet1);
      }
    },
    {
      label: 'Столбец',
      icon: 'pi pi-plus',
      command: () => {
        // toast.add({ severity: 'warn', summary: 'Search Completed', detail: 'No results found', life: 3000 });
        let maxCol = columns.value.length + 1;
        let colName = intToExcelCol(maxCol);
        workbook.Sheets.Sheet1[colName + 1] = {v:''};
        genlineItems(workbook.Sheets.Sheet1);
      }
    },
    {
      label: 'Столбец',
      icon: 'pi pi-minus',
      command: () => {
        // toast.add({ severity: 'warn', summary: 'Search Completed', detail: 'No results found', life: 3000 });
        let maxCol = columns.value.length;
        let colName = intToExcelCol(maxCol);
        let colName0;
        for(let key in workbook.Sheets.Sheet1){
          colName0 = key.replace(/[^a-zA-Z]+/g, '');
          if(colName0 == colName) delete workbook.Sheets.Sheet1[key];
        }
        genlineItems(workbook.Sheets.Sheet1);
      }
    }
  ]);
  const toggle = (event) => {
    menu.value.toggle(event);
  };
  const save = (event) => {
    props.modelValue.workbook = workbook;
    axios.patch(point+'/'+props.modelValue.id, props.modelValue)
      .then(response => toast.add({ severity: 'success', summary: 'Success', detail: 'Сохранено', life: 3000 }));
  };
</script>
