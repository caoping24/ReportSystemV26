<template>
  <div id="leaderDashboardMain">
    <!-- 1. 新增：标签栏 + 刷新按钮 容器（Flex布局） -->
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
            v-for="(page, index) in totalPages" 
            :key="page"
            :tab="tabNames[index]"
          />
        </a-tabs>
      </div>
      <!-- 新增：刷新按钮（居右） -->
      <a-tag 
        color="default"
        style="
            margin-left: 10px;
            font-size: 14px;
            padding: 3px 12px;
            border-radius: 6px;
            letter-spacing: 1px;
            background: #f5f7fa;
            border: 1px solid #e5e6eb;">
      {{ serverTime }}
      </a-tag>
    </div>
    <!-- 动态渲染的组件（原有逻辑不变） -->
    <component 
      :is="currentComponent" 
      ref="dashboardRef"
      :key="currentPage" 
      class="dashboard-content"
    />
  </div>
</template>

<script lang="ts" setup>
// 原有导入逻辑完全保留
import { ref, computed, defineAsyncComponent, watch, nextTick, onMounted, onUnmounted } from "vue";
import { message } from "ant-design-vue";
// 2026年5月14日新增：获取服务器时间的API导入
import { getServerTime } from "@/api/Dashboard";
const serverTime = ref("");
let serverNow: Date | null = null;
let timer: number | undefined;

function pad2(n: number) {
  return String(n).padStart(2, "0");
}
function format(d: Date) {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
}
onMounted(async () => {
  const res = await getServerTime(); // 你现在已经有这个 API
  const s = res.data.serverTime as string; // 按你的返回字段名调整

  serverNow = new Date(s.replace(" ", "T"));
  serverTime.value = format(serverNow);

  timer = window.setInterval(() => {
    if (!serverNow) return;
    serverNow = new Date(serverNow.getTime() + 1000);
    serverTime.value = format(serverNow);
  }, 1000);
});
onUnmounted(() => {
  if (timer) window.clearInterval(timer);
});
//
// 1. 原有异步导入子组件逻辑（完全不变）
const loadDashboardComponent = (page: number) => {
  return defineAsyncComponent(() => 
    import(`./LeaderDashboard/LeaderDashboard${page}.vue`)
  );
};

// 2. 分页状态管理（核心修改：标签选中状态持久化）
const totalPages = ref(3);
const currentPage = ref(1);

// 【新增】自定义标签名称数组（按顺序对应三个子页面）
const tabNames = ref([
  "配料 氨化 闪蒸",  // 对应 LeaderDashboard1
  "结晶",      // 对应 LeaderDashboard2
  "能耗"       // 对应 LeaderDashboard3
]);

// 核心修改1：将activeKey从computed改为ref直接管理，确保选中状态持久
const activeKey = ref<string>("1"); // 初始选中第一个标签

// 核心修改2：监听activeKey变化，同步更新currentPage（确保状态一致）
watch(activeKey, (newKey) => {
  const page = Number(newKey);
  if (page >= 1 && page <= totalPages.value) {
    currentPage.value = page;
  }
});

// 原有计算属性不变
const currentComponent = computed(() => {
  return loadDashboardComponent(currentPage.value);
});

// 核心修改3：调整标签切换逻辑，直接更新activeKey
const handleTabChange = (key: string) => {
  const page = Number(key);
  if (page < 1 || page > totalPages.value) return;
  activeKey.value = key; // 直接更新选中的标签key（持久化选中状态）
  message.success(`切换到${tabNames.value[page - 1]}`);
};

// 3. 原有子组件引用（完全不变）
const dashboardRef = ref<any>(null);

// ========== 仅新增：刷新按钮相关逻辑 ==========
const refreshLoading = ref(false); // 刷新按钮加载状态

// 调用子组件原有fetchAllData方法（核心：不修改子组件内部逻辑）
const handleRefresh = async () => {
  try {
    refreshLoading.value = true;
    const currentInstance = dashboardRef.value;
    
    if (!currentInstance) {
      message.warning("当前页面未加载完成，无法刷新");
      return;
    }
    // 调用子组件原有刷新方法（子组件仅暴露，不修改内部逻辑）
    if (typeof currentInstance.fetchAllData === "function") {
      await currentInstance.fetchAllData();
      message.success(`${tabNames.value[currentPage.value - 1]}刷新成功`);
    } else {
      message.error("当前页面无刷新方法");
    }
  } catch (error) {
    console.error("刷新失败：", error);
    message.error(`${tabNames.value[currentPage.value - 1]}刷新失败，请重试`);
  } finally {
    refreshLoading.value = false;
  }
};

// 监听标签切换，重置刷新按钮状态（可选优化）
watch(currentPage, () => {
  refreshLoading.value = false;
});
</script>

<style scoped>

/* 新增：标签栏+刷新按钮布局样式 */
.tabs-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  width: 100%;
  box-sizing: border-box;
  margin-bottom: 16px;
}
.tabs-container {
  width: calc(100% - 120px); /* 给刷新按钮留空间 */
  box-sizing: border-box;
}
.refresh-btn {
  margin-left: 16px;
  white-space: nowrap;
}
/* 原有样式完全保留 */
.dashboard-content {
  width: 100%;
  box-sizing: border-box;
}
/* 核心修改4：强化标签选中样式（和TableEditable.vue保持一致），确保选中状态持久且明显 */
:deep(.ant-tabs-card) {
  border: none;
  --ant-tabs-card-head-background: #f5f8fa;
  --ant-tabs-nav-item-active-color: #1890ff;
}
/* 强制生效的选中标签样式（关键：持久化视觉状态） */
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
/* 增强悬浮交互 */
:deep(.ant-tabs-tab:hover) {
  color: #096dd9 !important;
}
#leaderDashboardMain {
  width: 100%;
  box-sizing: border-box;
  padding: 16px;
  background-color: #f5f7fa;
  min-height: calc(100vh - 40px);
}
</style>