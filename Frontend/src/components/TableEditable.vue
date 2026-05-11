<template>
  <div id="tableEditableMain">
    <!-- 标签栏 + 日期选择器/按钮 容器 -->
    <div class="tabs-header">
      <div class="tabs-container">
        <a-tabs 
          v-model:activeKey="activeKey" 
          @change="handleTabChange"
          type="card"
          size="large"
          :tabBarStyle="{ marginBottom: 0 }" 
        >
          <a-tab-pane 
            v-for="page in totalPages" 
            :key="page"
            :tab="page === 1 ? '闪发器冷凝液' 
                : page === 2 ? '反应液' 
                : page === 3 ? '一次/二次结晶物/产品' 
                : page === 4 ? '一次母液' 
                : page === 5 ? '母液脱色前后/废液' 
                : page === 6 ? '能源消耗' 
                : `录入数据 ${page}`"
          />
        </a-tabs>
      </div>

      <!-- 日期选择器 + 查询按钮（删除了重载按钮） -->
      <div class="date-actions">
        <el-date-picker
          v-model="selectedDate"
          type="date"
          placeholder="选择查询日期"
          format="YYYY-MM-DD"
          value-format="YYYY-MM-DD"
          :disabled-date="disabledFutureDate"
          :picker-options="{
            shortcuts: [
              {
                text: '今天',
                onClick: () => {
                  selectedDate.value = new Date().toISOString().split('T')[0];
                },
              },
            ],
          }"
          :size="getComponentSize()"
          :style="getDatePickerStyle()"
        />
        <el-button
          type="primary"
          @click="handleQuery"
          :size="getComponentSize()"
        >
          查询
        </el-button>
      </div>
    </div>

    <!-- 动态渲染子组件：新增传递type参数 -->
    <component 
      :is="currentComponent" 
      ref="tableEditableRef"
      :key="activeKey"
      class="table-editable-content"
      :selected-date="selectedDate"
      :type="currentTab.type"
    />
  </div>
</template>

<script lang="ts" setup>
import { ref, computed, defineAsyncComponent, watch, onMounted, onUnmounted, nextTick } from "vue";
import { message } from "ant-design-vue";
import { ElMessage } from "element-plus";
import { ElDatePicker, ElButton } from "element-plus";

// 1. 异步导入子组件
const loadTableEditableComponent = (page: number) => {
  return defineAsyncComponent(() => 
    import(`./TablePage/TableEditable${page}.vue`)
  );
};

// 2. 分页状态管理 + 标签页与type的映射
interface EditableTab {
  key: string;
  tab: string;
  componentPage: number;
  type: number;
}

const tablePages: EditableTab[] = [
  { key: "1", tab: "检测数据 1", componentPage: 1, type: 1 },
  { key: "2", tab: "检测数据 2", componentPage: 2, type: 2 },
  { key: "3", tab: "检测数据 3", componentPage: 3, type: 3 },
  { key: "4", tab: "检测数据 4", componentPage: 4, type: 4 },
  { key: "5", tab: "检测数据 5", componentPage: 5, type: 5 },
  { key: "waste", tab: "废液", componentPage: 6, type: 7 },
  { key: "6", tab: "检测数据 6", componentPage: 6, type: 6 },
];

const activeKey = ref<string>("1");
const currentTab = computed(() => {
  return tablePages.find((item) => item.key === activeKey.value) || tablePages[0];
});

// 计算当前要渲染的子组件
const currentComponent = computed(() => {
  return loadTableEditableComponent(currentTab.value.componentPage);
});

// 处理标签切换事件
const handleTabChange = (key: string) => {
  const targetTab = tablePages.find((item) => item.key === key);
  if (!targetTab) return;
  activeKey.value = key; 
  message.success(`切换到${targetTab.tab}`);
};

// 3. 日期选择器核心逻辑
const selectedDate = ref<string>(new Date().toISOString().split('T')[0]);
const screenWidth = ref<number>(window.innerWidth);

const handleResize = () => {
  screenWidth.value = window.innerWidth;
};

const screenGrade = computed(() => {
  if (screenWidth.value < 1366) return "small";
  if (screenWidth.value < 1920) return "normal";
  return "large";
});

const getComponentSize = () => {
  return screenGrade.value === "small" ? "small" : "default";
};

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

const disabledFutureDate = (date: Date): boolean => {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const selectDate = new Date(date);
  selectDate.setHours(0, 0, 0, 0);
  return selectDate.getTime() > today.getTime();
};

// 4. 子组件引用
const tableEditableRef = ref<any>(null);

// 5. 查询按钮逻辑
const handleQuery = async () => {
  if (!selectedDate.value) {
    ElMessage.warning("请先选择查询日期");
    return;
  }

  try {
    const currentInstance = tableEditableRef.value;
    if (!currentInstance) return;

    if (typeof currentInstance.fetchTableData === "function") {
      await currentInstance.fetchTableData();
      ElMessage.success(`${currentTab.value.tab}【${selectedDate.value}】数据加载成功`);
    } else {
      ElMessage.error("当前表格无查询方法");
    }
  } catch (error) {
    console.error("表格查询失败：", error);
    ElMessage.error(`${currentTab.value.tab}查询失败，请重试`);
  }
};

// 6. 监听标签切换，自动查询当前日期数据
watch(activeKey, () => {
  nextTick(() => handleQuery());
});

// 7. 生命周期
onMounted(() => {
  window.addEventListener("resize", handleResize);
  handleQuery();
});

onUnmounted(() => {
  window.removeEventListener("resize", handleResize);
});
</script>

<!-- 样式部分保持不变 -->
<style scoped>
#tableEditableMain {
  width: 100%;
  box-sizing: border-box;
  padding: 16px;
  background-color: #f5f7fa;
  min-height: calc(100vh - 40px);
}

.tabs-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  width: 100%;
  box-sizing: border-box;
  margin-bottom: 16px;
}

.tabs-container {
  width: calc(100% - 380px);
  box-sizing: border-box;
}

.date-actions {
  display: flex;
  gap: 10px;
  align-items: center;
  flex-wrap: wrap;
  padding: 0 4px;
  white-space: nowrap;
}

.table-editable-content {
  width: 100%;
  box-sizing: border-box;
}

:deep(.ant-tabs-card) {
  border: none;
  --ant-tabs-card-head-background: #f5f8fa;
  --ant-tabs-nav-item-active-color: #1890ff;
}

:deep(.ant-tabs-tab-active) {
  background-color: #fff !important;
  color: #1890ff !important;
  font-weight: 700 !important;
  border-bottom: 2px solid #1890ff !important;
}

:deep(.ant-tabs-card .ant-tabs-tab-active) {
  border-color: #1890ff #1890ff #fff !important;
  box-shadow: 0 2px 4px rgba(24, 144, 255, 0.1) !important;
}

:deep(.ant-tabs-tab:hover) {
  color: #096dd9 !important;
}

@media screen and (max-width: 1366px) {
  .tabs-container {
    width: calc(100% - 320px);
  }
  .date-actions {
    gap: 8px;
  }
}
</style>
