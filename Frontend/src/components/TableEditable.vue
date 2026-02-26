<template>
  <a-tabs
    type="card"
    :style="{
      maxWidth: '100%',
      margin: '20px auto',
      padding: '0 10px',
    }"
    :tab-bar-gutter="getTabGutter()"
  >
    <a-tab-pane key="data-edit" tab="检测数据">
      <div class="table-container">
        <div class="date-selector">
          <el-date-picker
            v-model="selectedDate"
            type="date"
            placeholder="选择查询日期"
            @change="fetchTableData"
            format="YYYY-MM-DD"
            value-format="YYYY-MM-DD"
            :disabled-date="disabledFutureDate"
            :picker-options="{
              shortcuts: [
                {
                  text: '今天',
                  onClick: () => {
                    selectedDate.value = new Date().toISOString().split('T')[0];
                    fetchTableData();
                  },
                },
              ],
            }"
            :size="getComponentSize()"
            :style="getDatePickerStyle()"
          />
          <el-button
            type="primary"
            @click="fetchTableData"
            :size="getComponentSize()"
          >
            查询
          </el-button>
          <el-button
            type="primary"
            @click="reloadTableData"
            :size="getComponentSize()"
          >
            重载
          </el-button>
        </div>

        <!-- 关键修改1：重构表格容器，拆分表头和内容区域 -->
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
    </a-tab-pane>
    <a-tab-pane key="data-view" tab="数据预览">
      <div
        style="padding: 20px; text-align: center"
        :style="{ fontSize: getFontSize() }"
      >
        数据预览模块（可自定义内容）
      </div>
    </a-tab-pane>
  </a-tabs>
</template>

<script setup lang="ts">
// 【脚本部分完全不变】
import { ref, onMounted, computed, onUnmounted, nextTick, h } from "vue";
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

// 响应式数据
const selectedDate = ref<string>("");
const tableHeaders = ref<TableHeader[]>([]);
const tableData = ref<TableRow[]>([]);
const screenWidth = ref<number>(window.innerWidth);
const screenHeight = ref<number>(window.innerHeight); // 新增：监听屏幕高度

// 监听窗口大小变化（同时监听宽+高）
const handleResize = () => {
  screenWidth.value = window.innerWidth;
  screenHeight.value = window.innerHeight;
  // 窗口变化后重新计算布局，避免高度偏差
  nextTick(() => {
    document.documentElement.style.overflowY = "hidden"; // 确保网页级滚动条不出现
  });
};

// 屏幕分级（1080p属于normal）
const screenGrade = computed(() => {
  if (screenWidth.value < 1366) return "small";
  if (screenWidth.value < 1920) return "normal"; // 1080p(1920*1080)归为normal
  return "large";
});

// 组件尺寸计算
const getComponentSize = () => {
  return screenGrade.value === "small" ? "small" : "default";
};

const getTabGutter = () => {
  return screenGrade.value === "small" ? 8 : 16;
};

// 日期选择器宽度控制
const getDatePickerStyle = () => {
  const styles: Record<string, string> = {
    flexShrink: "0",
    padding: "0 4px",
  };
  switch (screenGrade.value) {
    case "small":
      styles.width = "180px";
      break;
    case "normal":
      styles.width = "200px";
      break;
    case "large":
      styles.width = "220px";
      break;
  }
  return styles;
};

// 表格列宽/输入框宽度/字体样式计算（保持不变）
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

// 业务逻辑保持不变
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

const fetchTableData = async (): Promise<void> => {
  if (!selectedDate.value) {
    ElMessage.warning("请先选择查询日期");
    return;
  }

  try {
    const res = await HourData({ date: selectedDate.value });
    const originData = res?.data || [];

    if (originData.length === 0) {
      tableData.value = [];
      ElMessage.info(`【${selectedDate.value}】暂无小时数据`);
      return;
    }

    const formatTableData = originData.map((item: HourDataItem) => {
      if (!item)
        return {
          hour: 0,
          date: selectedDate.value,
          isNextDay: false,
          cells: {},
        } as TableRow;
      const cellData = item.cells || {};
      return { ...item, ...cellData } as TableRow;
    });

    tableData.value = formatTableData;
    ElMessage.success(`【${selectedDate.value}】小时数据加载成功`);
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

const reloadTableData = async (): Promise<void> => {
  if (!selectedDate.value) {
    ElMessage.warning("请先选择查询日期");
    return;
  }

  try {
    ElMessage.info(`正在重载【${selectedDate.value}】数据，请稍候...`);
    const nextDay = new Date(selectedDate.value);
    nextDay.setDate(nextDay.getDate() + 1);
    const reloadParams: ReloadDataParams = {
      type: 1,
      time: nextDay.toISOString().split("T")[0],
    };
    await ReloadData(reloadParams);
    await fetchTableData();
    ElMessage.success(`【${selectedDate.value}】数据重载完成`);
  } catch (error) {
    ElMessage.error("数据重载失败，请重试");
    console.error("reloadTableData error:", error);
  }
};

// 初始化逻辑
onMounted(async () => {
  window.addEventListener("resize", handleResize);
  await fetchTableHeaders();
  selectedDate.value = new Date().toISOString().split("T")[0];
  await fetchTableData();

  // 初始化后强制隐藏网页级滚动条
  nextTick(() => {
    document.documentElement.style.overflowY = "hidden";
    document.body.style.overflowY = "hidden";
  });
});

onUnmounted(() => {
  window.removeEventListener("resize", handleResize);
  // 组件卸载后恢复网页滚动条
  document.documentElement.style.overflowY = "auto";
  document.body.style.overflowY = "auto";
});
</script>

<style scoped>
/* 核心修复：全局容器溢出控制 */
:deep(html),
:deep(body) {
  margin: 0;
  padding: 0;
  overflow-x: hidden; /* 禁止横向滚动 */
  overflow-y: hidden; /* 禁止网页级纵向滚动 */
  height: 100%;
}

:deep(.ant-tabs-card) {
  --ant-tabs-card-head-background: #f8f9fa;
  border-radius: 4px;
  width: 100%;
  height: calc(100vh - 40px); /* 适配页面上下边距，避免高度溢出 */
  overflow: hidden; /* 禁止tabs容器溢出 */
}

/* 日期选择器布局 */
.date-selector {
  margin-bottom: 10px;
  display: flex;
  gap: 10px;
  align-items: center;
  flex-wrap: wrap;
  padding: 0 4px;
  width: 100%;
}

/* 关键修改2：调整表格容器样式，适配固定表头 */
.table-scroll-wrapper {
  width: 100%;
  overflow: hidden; /* 隐藏容器溢出，仅表格内部滚动 */
  box-sizing: border-box;
  padding: 0 1px;
  margin: 4px 0;
  /* 高度计算：视口高度 - tabs头部 - 日期选择器 - 边距 */
  height: calc(100vh - 160px);
  scrollbar-gutter: stable;
}

/* 关键修改3：固定表头样式 + 表格内容滚动 */
:deep(.el-table) {
  --el-table-header-text-color: #333;
  --el-table-row-hover-bg-color: #f8f9fa;
   border: 1px solid #e6e6e6 !important;
}

/* 固定表头 */
:deep(.fixed-table-header) {
  position: sticky;
  top: 0;
  z-index: 10;
  background-color: #fff;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05); /* 可选：增加表头阴影，提升视觉区分 */
}

/* 表格内容区域滚动 */
:deep(.el-table__body-wrapper) {
  overflow-y: auto !important; /* 仅内容纵向滚动 */
  overflow-x: auto !important; /* 横向滚动（如需） */
  height: calc(
    100% - 40px
  ) !important; /* 表头高度约40px，内容区域自适应剩余高度 */
}

/* 不同屏幕的高度适配（修正1080p的高度偏差） */
@media screen and (max-width: 1366px) {
  .table-scroll-wrapper {
    height: calc(100vh - 150px);
  }
  .date-selector {
    gap: 8px;
  }
  :deep(.el-date-picker) {
    max-width: 180px !important;
  }
  /* 小屏适配表头高度 */
  :deep(.el-table__body-wrapper) {
    height: calc(100% - 36px) !important;
  }
}

/* 1080p（1920*1080）专属适配 */
@media screen and (min-width: 1367px) and (max-width: 1919px) {
  .table-scroll-wrapper {
    height: calc(100vh - 160px); /* 精准匹配1080p视口高度 */
  }
  /* 确保tabs容器高度不溢出 */
  :deep(.ant-tabs-card) {
    height: calc(100vh - 40px);
  }
}

@media screen and (min-width: 1920px) {
  .table-scroll-wrapper {
    height: calc(100vh - 170px);
  }
  /* 大屏适配表头高度 */
  :deep(.el-table__body-wrapper) {
    height: calc(100% - 44px) !important;
  }
}

/* 滚动条样式优化（仅表格内容区域显示滚动条） */
:deep(.el-table__body-wrapper::-webkit-scrollbar) {
  height: 12px;
  width: 8px;
}

:deep(.el-table__body-wrapper::-webkit-scrollbar-thumb) {
  background-color: #ccc;
  border-radius: 3px;
}

/* 表格样式 */
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
/* 日期选择器宽度限制 */
:deep(.el-date-picker) {
  max-width: 220px !important;
  width: 100% !important;
}

/* 其他样式保持不变 */
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
}
@media screen and (max-width: 1366px) {
  :deep(.el-table__fixed) {
    width: 50px !important;
  }
}
@media screen and (min-width: 1367px) and (max-width: 1919px) {
  :deep(.el-table__fixed) {
    width: 60px !important;
  }
}
@media screen and (min-width: 1920px) {
  :deep(.el-table__fixed) {
    width: 60px !important;
  }
}
</style>
