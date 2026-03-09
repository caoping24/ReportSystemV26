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

      <!-- 替换原有刷新按钮：日期选择器 + 查询 + 重载按钮 -->
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
        <el-button
          type="primary"
          @click="handleReload"
          :size="getComponentSize()"
        >
          重载
        </el-button>
      </div>
    </div>

    <!-- 动态渲染子组件：传递主组件的selectedDate -->
    <component 
      :is="currentComponent" 
      ref="tableEditableRef"
      :key="currentPage" 
      class="table-editable-content"
      :selected-date="selectedDate"
    />
  </div>
</template>

<script lang="ts" setup>
import { ref, computed, defineAsyncComponent, watch, onMounted, onUnmounted, nextTick } from "vue";
import { message } from "ant-design-vue";
import { ElMessage } from "element-plus";

// 导入Element Plus组件（全局注册可省略，非全局需导入）
import { ElDatePicker, ElButton } from "element-plus";

// 1. 异步导入子组件
const loadTableEditableComponent = (page: number) => {
  return defineAsyncComponent(() => 
    import(`./TablePage/TableEditable${page}.vue`)
  );
};

// 2. 分页状态管理（核心优化：确保选中状态持久化）
const totalPages = ref(2); // 总标签页数，可根据实际需求调整
const currentPage = ref(1);
const activeKey = ref<string>("1"); // 改为ref直接管理，初始选中第一个标签

// 监听activeKey变化，同步更新currentPage，确保状态一致
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
  activeKey.value = key; // 直接更新选中的标签key
  message.success(`切换到第 ${page} 个可编辑表格`);
};

// 3. 日期选择器核心逻辑
const selectedDate = ref<string>(new Date().toISOString().split('T')[0]);
const screenWidth = ref<number>(window.innerWidth);

// 监听窗口大小变化
const handleResize = () => {
  screenWidth.value = window.innerWidth;
};

// 屏幕尺寸分级
const screenGrade = computed(() => {
  if (screenWidth.value < 1366) return "small";
  if (screenWidth.value < 1920) return "normal";
  return "large";
});

// 动态计算组件尺寸
const getComponentSize = () => {
  return screenGrade.value === "small" ? "small" : "default";
};

// 动态设置日期选择器样式
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

// 禁用未来日期
const disabledFutureDate = (date: Date): boolean => {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const selectDate = new Date(date);
  selectDate.setHours(0, 0, 0, 0);
  return selectDate.getTime() > today.getTime();
};

// 4. 子组件引用
const tableEditableRef = ref<any>(null);

// 5. 查询按钮逻辑：调用子组件的查询方法
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

// 6. 重载按钮逻辑：调用子组件的重载方法
const handleReload = async () => {
  if (!selectedDate.value) {
    ElMessage.warning("请先选择查询日期");
    return;
  }

  try {
    const currentInstance = tableEditableRef.value;
    if (!currentInstance) {
      ElMessage.warning("当前表格未加载完成，无法重载");
      return;
    }

    if (typeof currentInstance.reloadTableData === "function") {
      await currentInstance.reloadTableData();
      ElMessage.success(`第 ${currentPage.value} 个表格【${selectedDate.value}】数据重载完成`);
    } else {
      ElMessage.error("当前表格无重载方法");
    }
  } catch (error) {
    console.error("表格重载失败：", error);
    ElMessage.error(`第 ${currentPage.value} 个表格重载失败，请重试`);
  }
};

// 7. 监听标签切换，自动查询当前日期数据
watch(currentPage, () => {
  nextTick(() => handleQuery());
});

// 8. 生命周期：监听窗口大小 + 初始化查询
onMounted(() => {
  window.addEventListener("resize", handleResize);
  handleQuery(); // 初始化时自动查询今日数据
});

onUnmounted(() => {
  window.removeEventListener("resize", handleResize);
});
</script>

<style scoped>
/* 主容器样式 */
#tableEditableMain {
  width: 100%;
  box-sizing: border-box;
  padding: 16px;
  background-color: #f5f7fa;
  min-height: calc(100vh - 40px);
}

/* 标签栏+日期按钮容器 */
.tabs-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  width: 100%;
  box-sizing: border-box;
  margin-bottom: 16px;
}

/* 标签容器 */
.tabs-container {
  width: calc(100% - 380px); /* 预留日期选择器+按钮宽度 */
  box-sizing: border-box;
}

/* 日期选择器+按钮容器 */
.date-actions {
  display: flex;
  gap: 10px;
  align-items: center;
  flex-wrap: wrap;
  padding: 0 4px;
  white-space: nowrap;
}

/* 子组件容器 */
.table-editable-content {
  width: 100%;
  box-sizing: border-box;
}

/* 核心：强化Antd标签选中样式，确保选中状态持久且明显 */
:deep(.ant-tabs-card) {
  border: none;
  --ant-tabs-card-head-background: #f5f8fa;
  --ant-tabs-nav-item-active-color: #1890ff;
}

/* 选中标签的核心样式（强制生效） */
:deep(.ant-tabs-tab-active) {
  background-color: #fff !important;
  color: #1890ff !important;
  font-weight: 700 !important;
  border-bottom: 2px solid #1890ff !important;
}

/* 卡片式标签选中时的边框样式 */
:deep(.ant-tabs-card .ant-tabs-tab-active) {
  border-color: #1890ff #1890ff #fff !important;
  box-shadow: 0 2px 4px rgba(24, 144, 255, 0.1) !important;
}

/* 标签悬浮样式（增强交互） */
:deep(.ant-tabs-tab:hover) {
  color: #096dd9 !important;
}

/* 小屏幕适配 */
@media screen and (max-width: 1366px) {
  .tabs-container {
    width: calc(100% - 320px);
  }
  .date-actions {
    gap: 8px;
  }
}
</style>