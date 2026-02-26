import { createRouter, createWebHistory, RouteRecordRaw } from "vue-router";
import { useLoginUserStore } from "@/store/useLoginUserStore";
import { message } from "ant-design-vue"; // 可选：提示信息
import { checkSessionTimeout } from "@/utils/sessionTimeout";
// 页面组件导入
import HomePage from "@/pages/HomePage.vue";
import UserManagePage from "@/pages/admin/UserManagePage.vue";
import UserLoginPage from "@/pages/user/UserLoginPage.vue";
import UserRegisterPage from "@/pages/user/UserRegisterPage.vue";
// 侧边栏布局组件
import BasicLayout from "@/layouts/BasicLayout.vue";
//领导仓
import LeaderDashboard from "@/components/LeaderDashboard.vue";
//手动table填写数据
import TableEditable from "@/components/TableEditable.vue";
// 路由规则
const routes: Array<RouteRecordRaw> = [
  // 根路径重定向到登录页
  {
    path: "/",
    redirect: "/user/login",
  },
  // 登录页（无布局）
  {
    path: "/user/login",
    name: "userLogin",
    component: UserLoginPage,
    meta: { requiresAuth: false },
  },
  // 注册页（无布局）
  {
    path: "/user/register",
    name: "userRegister",
    component: UserRegisterPage,
    meta: { requiresAuth: false },
  },
  // 登录后的主布局（带侧边栏）
  {
    path: "/app",
    component: BasicLayout,
    meta: { requiresAuth: true },
    children: [
      // 主页
      {
        path: "home",
        name: "homePage",
        component: HomePage,
      },
      // 用户管理页
      {
        path: "admin/userManage",
        name: "adminUserManage",
        component: UserManagePage,
      },
      {
        path: "components/leader-dashboard",
        name: "LeaderDashboard",
        component: LeaderDashboard,
        meta: { title: "数据看板" },
      },
      {
        path: "components/TableEditable",
        name: "TableEditable",
        component: TableEditable,
        meta: { title: "手动填写" },
      },
    ],
  },
  // 404页面
  {
    path: "/:pathMatch(.*)*",
    redirect: "/user/login",
  },
];

// 创建路由实例
const router = createRouter({
  history: createWebHistory(process.env.BASE_URL),
  routes,
});

// 路由守卫：登录校验
router.beforeEach((to, from, next) => {
  const loginUserStore = useLoginUserStore();
  // 从sessionStorage恢复登录状态（核心：刷新时恢复）
  loginUserStore.restoreLoginUser();

  // 新增：已登录但会话超时 → 强制登出
  const isLogin = !!loginUserStore.loginUser.id;
  if (isLogin && checkSessionTimeout()) {
    message.warning("登录状态已超时，请重新登录");
    loginUserStore.clearLoginUser();
    next({
      path: "/user/login",
      query: { redirect: encodeURIComponent(to.fullPath) },
    });
    return;
  }

  // 原有权限校验逻辑
  if (to.meta.requiresAuth) {
    if (isLogin) {
      next();
    } else {
      message.warning("请先登录后再访问");
      next({
        path: "/user/login",
        query: { redirect: encodeURIComponent(to.fullPath) },
      });
    }
  } else {
    if (
      isLogin &&
      (to.path === "/user/login" || to.path === "/user/register")
    ) {
      next("/app/home");
    } else {
      next();
    }
  }
});

export default router;
