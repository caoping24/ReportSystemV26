<template>
  <div class="layout-container">
    <!-- 左侧菜单栏 -->
    <aside class="sidebar">
      <!-- 侧边栏头部 -->
      <div class="sidebar-header">
        <img class="logo" src="../assets/hb.png" alt="logo" />
      </div>

      <!-- 侧边栏菜单 -->
      <a-menu
        v-model:selectedKeys="current"
        mode="vertical"
        :items="items"
        @click="doMenuClick"
        class="sidebar-menu"
      />

      <!-- 退出登录按钮 -->
      <div class="logout-btn">
        <a-button type="text" @click="handleLogout">退出登录</a-button>
      </div>
    </aside>

    <!-- 右侧主内容区 -->
    <main class="main-content">
      <router-view />
    </main>
  </div>
</template>

<script lang="ts" setup>
import { ref, h } from "vue";
import { MenuProps } from "ant-design-vue";
import { useRouter } from "vue-router";
import { useLoginUserStore } from "@/store/useLoginUserStore";
import { message } from "ant-design-vue";
// 1. 新增导入 EditOutlined 图标（可替换为你需要的任意图标）
import {
  HomeOutlined,
  DashboardOutlined,
  FileTextOutlined,
  EditOutlined,
} from "@ant-design/icons-vue";
import { userLogout } from "@/api/user";

const router = useRouter();
const loginUserStore = useLoginUserStore();

// 修复 TS 类型错误：使用 h 函数创建 VNode，避免类型推断问题
const items = ref<MenuProps["items"]>([
  {
    key: "/app/components/leader-dashboard",
    label: "数据看板",
    title: "数据看板",
    icon: () => h(DashboardOutlined),
  },
  {
    key: "/app/home",
    label: "报表下载",
    title: "报表下载",
    // 报表图标：FileTextOutlined
    icon: () => h(FileTextOutlined),
  },
  {
    key: "/app/components/TableEditable",
    label: "数据录入",
    title: "数据录入",
    // 2. 核心修改：将 DashboardOutlined 替换为 EditOutlined（手动填写的新图标）
    icon: () => h(EditOutlined),
  },
]);

// 菜单点击跳转
const doMenuClick = ({ key }: { key: string }) => {
  router.push(key);
};

// 同步路由选中状态
const current = ref<string[]>(["/app/home"]);
router.afterEach((to) => {
  current.value = [to.path];
});

// 退出登录逻辑（保持不变）
const handleLogout = async () => {
  try {
    await userLogout({});
    if (typeof loginUserStore.clearLoginUser === "function") {
      loginUserStore.clearLoginUser();
    } else {
      loginUserStore.loginUser = { userName: "未登录" };
    }
    localStorage.removeItem("loginUser");
    localStorage.removeItem("token");
    clearAllCookies();
    message.success("退出成功");
    router.replace({ path: "/user/login" });
  } catch (error) {
    console.error("退出登录接口调用失败：", error);
    if (typeof loginUserStore.clearLoginUser === "function") {
      loginUserStore.clearLoginUser();
    } else {
      loginUserStore.loginUser = { userName: "未登录" };
    }
    localStorage.removeItem("loginUser");
    localStorage.removeItem("token");
    clearAllCookies();
    message.success("已退出登录");
    router.replace({ path: "/user/login" });
  }
};

// 清空所有Cookie
const clearAllCookies = () => {
  const cookies = document.cookie.split(";");
  for (let i = 0; i < cookies.length; i++) {
    const cookie = cookies[i];
    const eqPos = cookie.indexOf("=");
    const name = eqPos > -1 ? cookie.substr(0, eqPos).trim() : cookie.trim();
    document.cookie = `${name}=; expires=Thu, 01 Jan 1970 00:00:00 GMT; path=/; domain=${window.location.hostname}`;
  }
};
</script>

<style scoped>
/* 样式部分保持不变 */
.layout-container {
  display: flex;

  height: 100vh;
  overflow: hidden;
}
.sidebar {
  width: 180px;
  height: 100%;
  background-color: #f5f9ff;
  color: #1e40af;
  box-shadow: 2px 0 12px rgba(0, 0, 0, 0.05);

  overflow-y: auto; /* 导航项过多时纵向滚动 */
  transition: width 0.3s; /* 折叠/展开动画 */
}
.sidebar-header {
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 0 12px;
  height: 60px;
  border-bottom: 1px solid #e0e7ff;
}
.logo {
  height: 30px;
  margin-right: 8px;
}
.title {
  color: #1e40af;
  font-size: 16px;
  font-weight: 600;
}
.sidebar-menu {
  border-right: none !important;
  padding: 12px 0;
  flex: 1;
}
.logout-btn {
  padding: 12px;
  border-top: 1px solid #e0e7ff;
}
.main-content {
  flex: 1;
  height: 100%;
  padding: 8px;
  background-color: #ffffff;
  overflow-y: auto;
}
:deep(.ant-menu-vertical .ant-menu-item) {
  height: 44px;
  line-height: 44px;
  margin: 0 8px !important;
  border-radius: 6px;
  padding-left: 20px !important;
  color: #1e40af !important;
  background-color: transparent !important;
}
:deep(.ant-menu-item-selected) {
  background-color: #3b82f6 !important;
  color: #ffffff !important;
}
:deep(.ant-menu-item:hover) {
  background-color: #dbeafe !important;
  color: #1e40af !important;
}
:deep(.ant-menu-title-content) {
  font-size: 14px;
  font-weight: 500;
}
:deep(.ant-btn-text) {
  color: #1e40af !important;
  display: block;
  width: 100%;
  text-align: center;
}
:deep(.ant-btn-text:hover) {
  background-color: #dbeafe !important;
  color: #1e40af !important;
}
</style>
