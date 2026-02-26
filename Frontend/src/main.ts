import { createApp } from "vue";
import App from "./App.vue";
import router from "./router";
import { createPinia } from "pinia";
import Antd from "ant-design-vue";
import "ant-design-vue/dist/reset.css";
import { message } from "ant-design-vue";
// 导入Store并恢复状态
import { useLoginUserStore } from "@/store/useLoginUserStore";
import ElementPlus from "element-plus";
import "element-plus/dist/index.css";
import zhCn from "element-plus/es/locale/lang/zh-cn";
const app = createApp(App);
app.use(ElementPlus, {
  locale: zhCn,
});
// 安装插件
app.use(createPinia());
app.use(router);
app.use(Antd);
app.use(ElementPlus);
// 1. 全局捕获未处理的 Promise 拒绝（核心逻辑）
window.addEventListener("unhandledrejection", (event) => {
  // 阻止默认行为（避免浏览器弹出错误弹窗）
  event.preventDefault();

  const reason = event.reason;
  // 过滤掉已经在 axios 拦截器中处理过的 401 错误
  if (
    reason?.response?.status === 401 ||
    reason?.response?.data?.code === 40100
  ) {
    console.info("401错误已在axios拦截器中处理，无需重复提示", reason);
    return;
  }

  // 对其他未捕获的错误，用 message 提示（这里直接使用即可）
  message.error("系统异常，请稍后重试");
  console.error("全局捕获未处理的Promise错误：", reason);
});

// 2. 正确注册 antd 插件（如果使用 antd 组件）
app.use(Antd); // 注册 antd 组件插件（可选，只用 message 则不需要）

// 恢复登录状态（双重保障）
const loginUserStore = useLoginUserStore();
loginUserStore.restoreLoginUser();

const originalResizeObserver = window.ResizeObserver;
window.ResizeObserver = class ResizeObserver extends originalResizeObserver {
  constructor(callback: ResizeObserverCallback) {
    // 用requestAnimationFrame做0ms延迟，避免循环通知
    super((entries, observer) => {
      requestAnimationFrame(() => {
        callback(entries, observer);
      });
    });
  }
};

app.mount("#app");
