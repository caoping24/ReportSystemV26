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
            :tab="`可编辑表格 ${page}`"
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
      :key="currentPage" 
      class="table-editable-content"
      :selected-date="selectedDate"
      :type="tabTypeMap[activeKey]"  
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
const totalPages = ref(6); 
const currentPage = ref(1);
const activeKey = ref<string>("1"); 
// 标签页key对应接口type参数（可根据实际业务调整）
const tabTypeMap = ref<Record<string, number>>({
  '1': 1,  // 第一个标签对应type=1
  '2': 2, 
  '3': 3,
  '4': 4,
  '5': 5,
  '6': 6
});

// 监听activeKey变化，同步更新currentPage
watch(activeKey, (newKey) => {
  const page = Number(newKey);
  if (page >= 1 && page <= totalPages.value) {
    currentPage.value = page;
  }
});

// 计算当前要渲染的子组件
const currentComponent = computed(() => {
  return loadTableEditableComponent(currentPage.value);
});

// 处理标签切换事件
const handleTabChange = (key: string) => {
  const page = Number(key);
  if (page < 1 || page > totalPages.value) return;
  activeKey.value = key; 
  message.success(`切换到第 ${page} 个可编辑表格`);
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
      ElMessage.success(`第 ${currentPage.value} 个表格【${selectedDate.value}】数据加载成功`);
    } else {
      ElMessage.error("当前表格无查询方法");
    }
  } catch (error) {
    console.error("表格查询失败：", error);
    ElMessage.error(`第 ${currentPage.value} 个表格查询失败，请重试`);
  }
};

// 6. 监听标签切换，自动查询当前日期数据
watch(currentPage, () => {
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