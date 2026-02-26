<template>
  <a-table
    :columns="columns"
    :data-source="dataSource"
    bordered
    :pagination="paginationConfig"
    :row-key="(record) => record.id"
  >
    <template #bodyCell="{ column, record, index }">
      <template v-if="column.key === 'index'">
        {{
          (paginationParams.pageIndex - 1) * paginationParams.pageSize +
          index +
          1
        }}
      </template>
      <template v-else-if="column.dataIndex === 'reportedTime'">
        {{ dayjs(record.reportedTime).format("YYYY-MM-DD") }}
      </template>
      <template v-else-if="column.dataIndex === 'createTime'">
        {{ dayjs(record.createdtime).format("YYYY-MM-DD HH:mm:ss") }}
      </template>

      <template v-else-if="column.key === 'action'">
        <a-button @click="emitRegenerate(record.reportedTime)">重建</a-button>
        <a-button @click="handleDownload(record.reportedTime)">下载</a-button>
      </template>
    </template>
  </a-table>
</template>

<script lang="ts" setup>
import dayjs from "dayjs";
import { defineProps, defineEmits } from "vue";
import type { TableProps } from "ant-design-vue/es/table";
import type { PaginationProps } from "ant-design-vue/es/pagination";

interface ReportTableProps {
  tabKey: string;
  columns: TableProps["columns"];
  dataSource: any[];
  paginationConfig: PaginationProps;
  paginationParams: {
    pageIndex: number;
    pageSize: number;
  };
}

const props = defineProps<ReportTableProps>();

const emit = defineEmits<{
  (e: "download", tabKey: string, reportedTime: string): void;
  (e: "regenerate", tabKey: string, reportedTime: string): void;
}>();

//格式化并触发 regenerate 事件
const emitRegenerate = (reportedTime: string | Date | undefined) => {
  if (!reportedTime) {
    emit("regenerate", props.tabKey, "");
    return;
  }
  const formattedTime = dayjs(reportedTime).format("YYYY-MM-DD") + " 09:00:01";
  emit("regenerate", props.tabKey, formattedTime);
};
//格式化并触发 download 事件
const handleDownload = (reportedTime: string | Date | undefined) => {
  if (!reportedTime) {
    emit("download", props.tabKey, "");
    return;
  }
  const formattedTime = dayjs(reportedTime).format("YYYY-MM-DD") + " 09:00:01";
  emit("download", props.tabKey, formattedTime);
};
</script>

<script lang="ts">
import { defineComponent } from "vue";

export default defineComponent({
  name: "ReportTable",
});
</script>

<style scoped>
/* 表格配色调整 - 表头背景改为纯白色 */
:deep(.ant-table) {
  --ant-table-header-text-color: #003399;
  /* 表头文字色（保留深蓝） */
  --ant-table-border-color: #e8f4fc;
  /* 表格边框色（浅蓝系） */
  --ant-table-row-hover-bg: #f0f8ff;
  /* 行hover背景（保留浅蓝） */
}

/* 表头样式 - 核心修改：背景改为纯白色 */
:deep(.ant-table-thead > tr > th) {
  background: #ffffff !important;
  /* 表头背景纯白 */
  color: #003399;
  /* 表头文字深蓝 */
  border-bottom: 2px solid #00aeef;
  /* 表头下边框（浅蓝飘带色） */
}

/* 下载按钮配色 */
:deep(.ant-table-cell .ant-btn) {
  color: #003399;
  border-color: #003399;
  background: #fff;
}

:deep(.ant-table-cell .ant-btn:hover) {
  color: #fff;
  background: #00aeef;
  border-color: #00aeef;
}

/* 表格行边框 */
:deep(.ant-table-tbody > tr > td) {
  border-bottom: 1px solid #e8f4fc;
  padding: 4px 16px !important;
}
</style>
