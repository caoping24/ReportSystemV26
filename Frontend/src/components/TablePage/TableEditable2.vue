<template>
  <div
    :style="{
      maxWidth: '100%',
      margin: '0 auto',
      padding: '0 10px',
      height: '100%',
    }"
  >
    <div class="table-container">
      <!-- 移除原有日期选择器/按钮区域 -->

      <!-- 调整表格容器高度（移除日期选择器后减少 marginTop） -->
      <div class="table-scroll-wrapper">
        <el-table
          :data="tableData"
          border
          style="width: 100%; table-layout: fixed"
          :cell-class-name="cellClassName"
          :size="getComponentSize()"
          empty-text="当前日期暂无小时数据"
          :header-cell-style="getHeaderCellStyle()"
          :cell-style="getCellStyle()"
          height="100%"
          :header-row-class-name="'fixed-table-header'"
        >
          <el-table-column
            v-for="(header, index) in tableHeaders"
            :key="index"
            :prop="header.prop"
            :label="header.label"
            :width="getColumnWidth(header.prop)"
            align="center"
            :show-overflow-tooltip="true"
            :fixed="header.prop === 'hour' ? 'left' : false"
          >
            <template #default="scope">
              <template v-if="header.prop === 'hour'">
                {{ scope.row[header.prop] }}
              </template>
              <template v-else>
                <template v-if="isCellDisabled(scope.row)">
                  {{ scope.row[header.prop] || "-" }}
                </template>
                <template v-else>
                  <el-input
                    v-model="scope.row[header.prop]"
                    :size="getComponentSize()"
                    @blur="handleCellEdit(scope.row, header.prop)"
                    :disabled="isCellDisabled(scope.row)"
                    maxlength="8"
                    :style="{ width: '100%', height: '100%' }"
                  />
                </template>
              </template>
            </template>
          </el-table-column>
        </el-table>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, computed, onUnmounted, nextTick, defineProps } from "vue";
import { ElMessage } from "element-plus";
import { Headers, HourData, SaveCell, ReloadData } from "@/api/TableEdit";

// 类型定义保持不变
interface TableHeader {
  prop: string;
  label: string;
}

interface HourDataItem {
  hour: number;
  date: string;
  isNextDay: boolean;
  cells?: Record<string, string>;
}

interface TableRow extends HourDataItem {
  [key: string]: any;
}

interface ReloadDataParams {
  type: number;
  time: string;
}

// 接收主组件传递的日期
const props = defineProps({
  selectedDate: {
    type: String,
    required: true,
    default: () => new Date().toISOString().split('T')[0]
  }
});

// 响应式数据：移除自身的selectedDate，改用props.selectedDate
const tableHeaders = ref<TableHeader[]>([]);
const tableData = ref<TableRow[]>([]);
const screenWidth = ref<number>(window.innerWidth);
const screenHeight = ref<number>(window.innerHeight);

// 保留原有逻辑，仅替换selectedDate为props.selectedDate
const handleResize = () => {
  screenWidth.value = window.innerWidth;
  screenHeight.value = window.innerHeight;
  nextTick(() => {
    document.documentElement.style.overflowY = "hidden";
  });
};

const screenGrade = computed(() => {
  if (screenWidth.value < 1366) return "small";
  if (screenWidth.value < 1920) return "normal";
  return "large";
});

const getComponentSize = () => {
  return screenGrade.value === "small" ? "small" : "default";
};

const getTabGutter = () => {
  return screenGrade.value === "small" ? 8 : 16;
};

// 移除getDatePickerStyle（主组件已实现）

const getColumnWidth = (prop: string) => {
  if (prop === "hour") {
    return screenGrade.value === "small" ? 50 : 60;
  }
  return screenGrade.value === "small"
    ? 80
    : screenGrade.value === "large"
    ? 100
    : 90;
};

const getInputWidth = () => {
  return screenGrade.value === "small"
    ? "70px"
    : screenGrade.value === "large"
    ? "90px"
    : "80px";
};

const getHeaderCellStyle = () => {
  const fontSize =
    screenGrade.value === "small"
      ? "11px"
      : screenGrade.value === "large"
      ? "13px"
      : "12px";
  return {
    fontSize,
    padding: "2px 0",
  };
};

const getCellStyle = () => {
  const fontSize =
    screenGrade.value === "small"
      ? "10px"
      : screenGrade.value === "large"
      ? "14px"
      : "13px";
  return {
    fontSize,
    padding: "2px 0",
  };
};

const getFontSize = () => {
  return screenGrade.value === "small"
    ? "12px"
    : screenGrade.value === "large"
    ? "14px"
    : "13px";
};

const disabledFutureDate = (date: Date): boolean => {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const selectDate = new Date(date);
  selectDate.setHours(0, 0, 0, 0);
  return selectDate.getTime() > today.getTime();
};

const isCellDisabled = (row: TableRow): boolean => {
  if (!row.date || row.hour === undefined || row.hour === null) return true;
  return row.isNextDay === true;
};

const cellClassName = ({
  row,
  column,
}: {
  row: TableRow;
  column: any;
}): string => {
  if (column.prop === "hour") return "disabled-cell";
  return isCellDisabled(row) ? "disabled-cell" : "";
};

const fetchTableHeaders = async (): Promise<void> => {
  try {
    const res = await Headers();
    if (res?.data) {
      tableHeaders.value = res.data;
      const hourHeader = tableHeaders.value.find(
        (item) => item.prop === "hour"
      );
      if (hourHeader) {
        tableHeaders.value = [
          hourHeader,
          ...tableHeaders.value.filter((item) => item.prop !== "hour"),
        ];
      }
    }
  } catch (error) {
    ElMessage.error("获取表格表头失败，请刷新页面");
    console.error("fetchTableHeaders error:", error);
  }
};

// 核心修改：使用props.selectedDate替代自身的selectedDate
const fetchTableData = async (): Promise<void> => {
  if (!props.selectedDate) {
    ElMessage.warning("请先选择查询日期");
    return;
  }

  try {
    const res = await HourData({ date: props.selectedDate });
    const originData = res?.data || [];

    if (originData.length === 0) {
      tableData.value = [];
      ElMessage.info(`【${props.selectedDate}】暂无小时数据`);
      return;
    }

    const formatTableData = originData.map((item: HourDataItem) => {
      if (!item)
        return {
          hour: 0,
          date: props.selectedDate,
          isNextDay: false,
          cells: {},
        } as TableRow;
      const cellData = item.cells || {};
      return { ...item, ...cellData } as TableRow;
    });

    tableData.value = formatTableData;
    ElMessage.success(`【${props.selectedDate}】小时数据加载成功`);
  } catch (error) {
    ElMessage.error("小时数据加载失败，请重试");
    console.error("fetchTableData error:", error);
  }
};

const handleCellEdit = async (row: TableRow, prop: string): Promise<void> => {
  if (prop === "hour" || isCellDisabled(row)) return;

  const saveParams = {
    date: row.date,
    hour: row.hour,
    prop: prop,
    value: row[prop] || "",
  };

  try {
    await SaveCell(saveParams);
    ElMessage.success(`已保存：${row.date} ${row.hour}点 - ${prop} 字段`);
  } catch (error) {
    ElMessage.error("单元格数据保存失败，请重试");
    console.error("handleCellEdit error:", error);
  }
};

// 核心修改：使用props.selectedDate替代自身的selectedDate
const reloadTableData = async (): Promise<void> => {
  if (!props.selectedDate) {
    ElMessage.warning("请先选择查询日期");
    return;
  }

  try {
    ElMessage.info(`正在重载【${props.selectedDate}】数据，请稍候...`);
    const nextDay = new Date(props.selectedDate);
    nextDay.setDate(nextDay.getDate() + 1);
    const reloadParams: ReloadDataParams = {
      type: 1,
      time: nextDay.toISOString().split("T")[0],
    };
    await ReloadData(reloadParams);
    await fetchTableData();
    ElMessage.success(`【${props.selectedDate}】数据重载完成`);
  } catch (error) {
    ElMessage.error("数据重载失败，请重试");
    console.error("reloadTableData error:", error);
  }
};

// 初始化逻辑：移除自身selectedDate的赋值，改为监听props变化
onMounted(async () => {
  window.addEventListener("resize", handleResize);
  await fetchTableHeaders();
  // 初始化时调用查询（使用主组件传递的日期）
  if (props.selectedDate) {
    await fetchTableData();
  }

  nextTick(() => {
    document.documentElement.style.overflowY = "hidden";
    document.body.style.overflowY = "hidden";
  });
});

// 监听主组件传递的日期变化，自动刷新数据
onUnmounted(() => {
  window.removeEventListener("resize", handleResize);
  document.documentElement.style.overflowY = "auto";
  document.body.style.overflowY = "auto";
});

// 暴露方法给主组件调用
defineExpose({
  fetchTableData,
  reloadTableData,
});
</script>

<style scoped>
/* 核心修改：调整表格容器高度（移除日期选择器后减少顶部间距） */
.table-scroll-wrapper {
  width: 100%;
  overflow: hidden;
  box-sizing: border-box;
  padding: 0 1px;
  margin: 0; /* 移除原有margin-top */
  /* 调整高度：移除日期选择器后减少偏移量 */
  height: calc(100vh - 80px);
  scrollbar-gutter: stable;
}

/* 其他样式保持不变，仅调整媒体查询中的高度 */
:deep(html),
:deep(body) {
  margin: 0;
  padding: 0;
  overflow-x: hidden;
  overflow-y: hidden;
  height: 100%;
}

:deep(.el-table) {
  --el-table-header-text-color: #333;
  --el-table-row-hover-bg-color: #f8f9fa;
  border: 1px solid #e6e6e6 !important;
}

:deep(.fixed-table-header) {
  position: sticky;
  top: 0;
  z-index: 10;
  background-color: #fff;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
}

:deep(.el-table__body-wrapper) {
  overflow-y: auto !important;
  overflow-x: auto !important;
  height: calc(100% - 40px) !important;
}

/* 小屏适配 */
@media screen and (max-width: 1366px) {
  .table-scroll-wrapper {
    height: calc(100vh - 70px);
  }
  :deep(.el-table__body-wrapper) {
    height: calc(100% - 36px) !important;
  }
}

/* 1080p适配 */
@media screen and (min-width: 1367px) and (max-width: 1919px) {
  .table-scroll-wrapper {
    height: calc(100vh - 80px);
  }
}

/* 大屏适配 */
@media screen and (min-width: 1920px) {
  .table-scroll-wrapper {
    height: calc(100vh - 90px);
  }
  :deep(.el-table__body-wrapper) {
    height: calc(100% - 44px) !important;
  }
}

/* 其余样式保持不变 */
:deep(.el-table__body-wrapper::-webkit-scrollbar) {
  height: 12px;
  width: 8px;
}

:deep(.el-table__body-wrapper::-webkit-scrollbar-thumb) {
  background-color: #ccc;
  border-radius: 3px;
}

:deep(.el-table td),
:deep(.el-table th) {
  padding: 2px 0 !important;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
:deep(.el-table .cell) {
  padding: 1px 1px !important;
}

.disabled-cell {
  background-color: #f5f5f5;
  color: #999;
  cursor: not-allowed;
}

:deep(.el-input__wrapper) {
  padding: 0 5px !important;
  box-sizing: border-box;
}

:deep(.el-picker-panel__content .el-date-table td.disabled) {
  color: #ccc !important;
  cursor: not-allowed !important;
}
:deep(.el-table__fixed) {
  z-index: 11 !important;
  background-color: #fff;
  border-right: 1px solid #e6e6e6 !important;
}
:deep(.el-table__fixed-header-wrapper) {
  z-index: 12 !important;
  background-color: #fff;
}
:deep(.fixed-table-header th.el-table__cell) {
  border-right: 1px solid #e6e6e6 !important;
}
:deep(.el-table__fixed th.el-table__cell) {
  border-bottom: 1px solid #e6e6e6 !important;
}
:deep(.el-table) {
  --el-table-header-text-color: #333;
  --el-table-row-hover-bg-color: #f8f9fa;
  border: 1px solid #e6e6e6 !important;
  --el-table-header-border-color: #e6e6e6;
  --el-table-border-color: #e6e6e6;
}
:deep(.el-table__fixed td),
:deep(.el-table__fixed th) {
  border-left: 1px solid #e6e6e6 !important;
}
@media screen and (max-width: 1366px) {
  :deep(.el-table th .cell) {
    font-weight: 500;
  }
  :deep(.el-input__wrapper) {
    font-size: 12px;
  }
  :deep(.el-table__fixed) {
    width: 50px !important;
  }
}

@media screen and (min-width: 1920px) {
  :deep(.el-table th .cell) {
    font-size: 14px;
    font-weight: 600;
  }
  :deep(.el-table td .cell) {
    font-size: 14px;
  }
  :deep(.el-input__wrapper) {
    font-size: 14px;
  }
  :deep(.el-table__fixed) {
    width: 60px !important;
  }
}
@media screen and (min-width: 1367px) and (max-width: 1919px) {
  :deep(.el-table__fixed) {
    width: 60px !important;
  }
}
</style>