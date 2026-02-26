<template>
  <div id="userManagePage">
    <a-tabs default-active-key="1" type="card" @change="handleTabChange">
      <!-- 原有4个报表标签页（简化版） -->
      <a-tab-pane v-for="item in reportTabs" :key="item.key" :tab="item.tab">
        <report-table
          :tab-key="item.key"
          :columns="columns"
          :data-source="tableData[item.key]?.list || []"
          :pagination-config="paginationConfig"
          :pagination-params="paginationParams"
          @download="downloadExcel"
          @regenerate="handleRegenerateRow"
        />
      </a-tab-pane>

      <!-- 批量下载报表标签页 -->
      <a-tab-pane tab="批量下载报表" key="5">
        <div class="batch-download-container">
          <!-- 标题和说明 -->
          <div class="batch-download-header">
            <h3 class="batch-download-title">批量报表下载</h3>
            <p class="batch-download-desc">
              根据选择的报表类型和月份，下载对应月份的报表文件
            </p>
          </div>

          <!-- 修改点1：添加ConfigProvider包裹表单，配置中文语言 -->
          <ConfigProvider :locale="zhCN">
            <!-- 表单区域 -->
            <a-form
              class="batch-download-form"
              layout="vertical"
              :label-col="{ span: 4 }"
              :wrapper-col="{ span: 20 }"
            >
              <a-form-item
                label="报表类型"
                :validate-status="!batchReportType ? 'error' : ''"
                :help="!batchReportType ? '请选择报表类型' : ''"
              >
                <a-select
                  v-model:value="batchReportType"
                  style="width: 100%; max-width: 300px"
                  placeholder="请选择报表类型"
                  allow-clear
                  @change="handleReportTypeChange"
                >
                  <a-select-option
                    v-for="item in reportTabs.filter(
                      (item) => item.key !== '4'
                    )"
                    :key="item.key"
                    :value="item.key"
                  >
                    {{ item.tab }}
                  </a-select-option>
                </a-select>
              </a-form-item>

              <!-- 修改点2：替换为单个月份选择框 -->
              <a-form-item
                label="选择月份"
                :validate-status="!batchMonth && batchReportType ? 'error' : ''"
                :help="!batchMonth && batchReportType ? '请选择报表月份' : ''"
              >
                <a-date-picker
                  v-model:value="batchMonth"
                  picker="month"
                  style="width: 100%; max-width: 300px"
                  placeholder="选择报表月份"
                  format="YYYY年MM月"
                  allow-clear
                  :disabled-date="disabledFutureDate"
                />
              </a-form-item>

              <!-- 操作按钮区域 -->
              <a-form-item :wrapper-col="{ offset: 4 }">
                <a-button
                  type="primary"
                  @click="batchDownloadZip"
                  :loading="isBatchDownloading"
                  class="batch-download-btn"
                >
                  <template #icon><DownloadOutlined /></template>
                  下载报表
                </a-button>

                <a-button @click="resetBatchForm" style="margin-left: 12px">
                  重置
                </a-button>

                <div class="batch-tips">
                  <InfoCircleOutlined style="margin-right: 4px" />
                  提示：下载所选月份的对应类型报表文件（默认按当月01日查询）
                </div>
              </a-form-item>
            </a-form>
          </ConfigProvider>
        </div>
      </a-tab-pane>
    </a-tabs>
  </div>
</template>

<script lang="ts" setup>
// script部分代码完全不变
import {
  getReportByPage,
  downloadReport,
  batchDownloadReportZip,
  regenerateReports,
} from "@/api/user";
import { message } from "ant-design-vue";
import { ref, reactive, computed } from "vue";
import dayjs from "dayjs";
import { DownloadOutlined, InfoCircleOutlined } from "@ant-design/icons-vue";
import ReportTable from "@/components/ReportTable.vue";

import { ConfigProvider } from "ant-design-vue";
import zhCN from "ant-design-vue/es/locale/zh_CN";
import "dayjs/locale/zh-cn";
dayjs.locale("zh-cn");

// ===================== 类型定义（精简整合） =====================
interface ReportItem {
  id: string;
  createdtime: string | Date;
}

interface TableDataItem {
  list: ReportItem[];
}

interface ReportTabItem {
  key: string;
  tab: string;
}

// ===================== 常量定义 =====================
const reportTabs: ReportTabItem[] = [
  { key: "1", tab: "日报表" },
  { key: "2", tab: "周报表" },
  { key: "3", tab: "月报表" },
  { key: "4", tab: "年报表" },
];

const columns = [
  { title: "序号", key: "index", width: 80, align: "center" },
  { title: "报表日期", dataIndex: "reportedTime", key: "reportedTime" },
  { title: "创建时间", dataIndex: "createTime", key: "createTime" },
  { title: "操作", key: "action", width: 120, align: "center" },
];

// ===================== 状态管理 =====================
const activeTabKey = ref("1");
const isFetching = ref(false);

const paginationParams = reactive({
  pageIndex: 1,
  pageSize: 10,
  total: 0,
});

const tableData: Record<string, TableDataItem> = reactive({
  "1": { list: [] },
  "2": { list: [] },
  "3": { list: [] },
  "4": { list: [] },
});

const batchReportType = ref<string>("");
// 修改点3：替换开始/结束时间为单个月份选择
const batchMonth = ref<dayjs.Dayjs | null>(null);
const isBatchDownloading = ref<boolean>(false);
// 新增：重新生成状态（避免并发）
const isRegenerating = ref<boolean>(false);

// ===================== 计算属性（简化年报表相关逻辑） =====================
const typeMap = {
  "3": "month",
} as Record<string, "date" | "month">;

const formatMap = {
  "3": "YYYY年MM月",
} as Record<string, string>;

const paginationConfig = computed(() => ({
  current: paginationParams.pageIndex,
  pageSize: paginationParams.pageSize,
  total: paginationParams.total,
  pageSizeOptions: ["10", "20", "50", "100"],
  showSizeChanger: true,
  showQuickJumper: true,
  showTotal: (total: number) => `共 ${total} 条记录`,
  onChange: (page: number, pageSize: number) => {
    paginationParams.pageIndex = page;
    paginationParams.pageSize = pageSize;
    reportTabs.map((item) => item.key).includes(activeTabKey.value) &&
      fetchData(activeTabKey.value);
  },
  onShowSizeChange: (current: number, size: number) => {
    paginationParams.pageIndex = 1;
    paginationParams.pageSize = size;
    reportTabs.map((item) => item.key).includes(activeTabKey.value) &&
      fetchData(activeTabKey.value);
  },
}));

// ===================== 方法定义（封装精简） =====================
const disabledFutureDate = (current: dayjs.Dayjs) => {
  return current?.isAfter(dayjs().endOf("month")) || false; // 调整为禁用未来月份
};

// 修改点4：更新表单验证逻辑
const validateBatchParams = () => {
  if (!batchReportType.value) {
    message.warning("请选择报表类型");
    return false;
  }
  if (!batchMonth.value) {
    message.warning("请选择报表月份");
    return false;
  }
  return true;
};

const handleFileDownload = (
  res: any,
  defaultFileName: string,
  fileType: "xlsx" | "zip"
) => {
  const typeMap = {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    zip: "application/zip",
  };

  const blob = new Blob([res.data], { type: typeMap[fileType] });
  const contentDisposition = res.headers?.["content-disposition"];
  let fileName = defaultFileName;

  if (contentDisposition) {
    const utf8Match = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
    if (utf8Match?.[1]) {
      fileName = decodeURIComponent(utf8Match[1]);
    } else {
      const normalMatch = contentDisposition.match(/filename=([^;]+)/i);
      if (normalMatch?.[1]) {
        fileName = decodeURIComponent(normalMatch[1].replace(/['"]/g, ""));
      }
    }
  }

  const url = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();

  document.body.removeChild(link);
  window.URL.revokeObjectURL(url);

  message.success(`${fileType === "xlsx" ? "报表" : "ZIP包"}下载成功`);
};

const downloadExcel = async (tabKey: string, reportedTime: string) => {
  if (!reportedTime) return message.warning("ID 不能为空");
  try {
    const res = await downloadReport(reportedTime, Number(tabKey));
    // 从响应头解析文件名
    const contentDisposition = res.headers?.["content-disposition"];
    let fileName = `报表_${reportedTime.substring(0, 10)}.xlsx`;
    if (contentDisposition) {
      const utf8Match = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
      if (utf8Match?.[1]) {
        fileName = decodeURIComponent(utf8Match[1]);
      } else {
        const normalMatch = contentDisposition.match(/filename=([^;]+)/i);
        if (normalMatch?.[1]) {
          fileName = decodeURIComponent(normalMatch[1].replace(/['"]/g, ""));
        }
      }
    }
    handleFileDownload(res, fileName, "xlsx");
  } catch (error) {
    console.error("报表下载失败：", error);
    message.error("下载失败：网络异常或接口错误");
  }
};
// 修改点5：更新批量下载参数组装逻辑
const batchDownloadZip = async () => {
  if (!validateBatchParams()) return;

  // 将选中的月份补全为 yyyy-mm-01 格式
  const selectedTime = batchMonth.value!.format("YYYY-MM-01");

  const params = {
    type: Number(batchReportType.value),
    timeStr: selectedTime, // 参数名改为time，格式为yyyy-mm-dd
  };

  try {
    isBatchDownloading.value = true;
    const res = await batchDownloadReportZip(params);
    // 根据实际返回的文件类型调整，这里假设还是zip，也可以改为xlsx
    handleFileDownload(
      res,
      `报表_${batchMonth.value!.format("YYYYMM")}.zip`,
      "zip"
    );
  } catch (error) {
    console.error("批量下载ZIP失败：", error);
    message.error("批量下载失败：网络异常或接口错误");
  } finally {
    isBatchDownloading.value = false;
  }
};

const handleTabChange = (key: string) => {
  activeTabKey.value = key;
  if (key !== "5") {
    paginationParams.pageIndex = 1;
    fetchData(key);
  }
};

// 修改点6：更新报表类型变更处理
const handleReportTypeChange = () => {
  batchMonth.value = null; // 清空月份选择
};

// 修改点7：更新重置表单逻辑
const resetBatchForm = () => {
  batchReportType.value = "";
  batchMonth.value = null;
};

// 新增：处理表格行的重新生成请求
const handleRegenerateRow = async (tabKey: string, reportedTime: string) => {
  if (!tabKey || !reportedTime) {
    return message.warning("参数缺失，无法重新生成");
  }
  try {
    isRegenerating.value = true;
    // 生成后端可直接解析的ISO 8601格式
    const timeStr = reportedTime.replace(" ", "T") + ".000";
    await regenerateReports({
      type: Number(tabKey),
      time: timeStr,
    });
    message.success("已提交重新生成请求，任务完成后请下载或刷新查看");
  } catch (error) {
    console.error("重新生成失败：", error);
    message.error("重新生成失败：网络异常或接口错误");
  } finally {
    isRegenerating.value = false;
  }
};

const fetchData = async (tabKey: string) => {
  if (isFetching.value) return;

  try {
    isFetching.value = true;
    const res = await getReportByPage({
      pageIndex: paginationParams.pageIndex,
      pageSize: paginationParams.pageSize,
      Type: Number(tabKey),
    });

    if (res.data) {
      tableData[tabKey].list = res.data.data || [];
      paginationParams.total = res.data.totalCount || 0;
    } else {
      message.error(`获取${tabKey}号表格数据失败`);
    }
  } catch (error) {
    console.error(`获取${tabKey}号表格数据异常：`, error);
    message.error(`获取${tabKey}号表格数据失败：网络异常`);
  } finally {
    isFetching.value = false;
  }
};

fetchData("1");
</script>

<style scoped>
/* 基础布局样式保留，仅删除标签页自定义配色 */
#userManagePage {
  padding: 20px;
}

/* 恢复标签页原有默认样式（删除之前的自定义标签页样式） */
.ant-tabs-card > .ant-tabs-nav .ant-tabs-tab {
  padding: 12px 24px;
}

/* ========== 仅保留非标签页的配色修改 ========== */
/* 1. 按钮配色 */
/* 主按钮（批量下载）- HEBANG主深蓝色 */
:deep(.ant-btn-primary) {
  background: #003399;
  border-color: #003399;
}
:deep(.ant-btn-primary:hover),
:deep(.ant-btn-primary:focus) {
  background: #0066cc; /* 深蓝hover色 */
  border-color: #0066cc;
}
/* 普通按钮（重置）- 浅蓝辅助色 */
:deep(.ant-btn-default) {
  color: #003399;
  border-color: #003399;
}
:deep(.ant-btn-default:hover) {
  color: #00aeef;
  border-color: #00aeef;
  background: #f0f8ff;
}

/* 2. 表单/选择器配色 */
/* 选择器/日期选择器激活边框 */
:deep(
    .ant-select-focused:not(.ant-select-disabled).ant-select:not(
        .ant-select-customize-input
      )
      .ant-select-selector,

  )
  :deep(.ant-picker-focused) {
  border-color: #00aeef !important;
  box-shadow: 0 0 0 2px rgba(0, 174, 239, 0.2);
}
/* 选择器hover */
:deep(.ant-select-selector:hover),
:deep(.ant-picker:hover) {
  border-color: #00aeef;
}

/* 3. 提示文字/图标配色 */
.batch-tips {
  margin-top: 12px;
  font-size: 12px;
  color: #666666;
  display: flex;
  align-items: center;
}
.batch-tips .anticon-infocircle {
  color: #00aeef; /* 浅蓝飘带色 */
}

/* 4. 批量下载容器样式（调整边框/阴影配色） */
.batch-download-container {
  margin-top: 16px;
  padding: 24px;
  background: #ffffff;
  border-radius: 12px;
  box-shadow: 0 2px 12px 0 rgba(0, 51, 153, 0.08); /* 主深蓝色浅阴影 */
  border: 1px solid #e8f4fc; /* 浅蓝系边框 */
}

.batch-download-header {
  margin-bottom: 24px;
  padding-bottom: 16px;
  border-bottom: 1px solid #e8f4fc; /* 浅蓝系分隔线 */
}

.batch-download-title {
  margin: 0 0 8px 0;
  font-size: 18px;
  font-weight: 600;
  color: #003399; /* 主深蓝色标题 */
}

.batch-download-desc {
  margin: 0;
  color: #666666;
  font-size: 14px;
  line-height: 1.5;
}

.batch-download-form {
  max-width: 800px;
}

/* 修改点8：移除日期范围组样式，调整表单间距 */
.batch-download-form .ant-form-item {
  margin-bottom: 16px !important;
}

.batch-download-btn {
  height: 40px;
  padding: 0 24px;
}

/* 5. 分页组件配色 */
:deep(.ant-pagination-item-active) {
  border-color: #003399;
  background: #003399;
}
:deep(.ant-pagination-item-active a) {
  color: #fff;
}
:deep(.ant-pagination-item:hover) {
  border-color: #00aeef;
}
:deep(.ant-pagination-item a:hover) {
  color: #00aeef;
}

/* 响应式适配 */
@media (max-width: 768px) {
  .batch-download-container {
    padding: 16px;
  }

  .batch-download-form .ant-form-item {
    min-width: 100%;
  }
}
</style>
