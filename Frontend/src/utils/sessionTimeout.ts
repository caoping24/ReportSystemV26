import { useLoginUserStore } from "@/store/useLoginUserStore";
import router from "@/router";
import { message } from "ant-design-vue";

// 超时时间（毫秒），按需调整：如30分钟 = 30 * 60 * 1000
export const SESSION_TIMEOUT = 30 * 60 * 1000;
// 优先从sessionStorage恢复最后操作时间（刷新保留）
let lastOperateTime = parseInt(
  sessionStorage.getItem("lastOperateTime") || `${Date.now()}`,
  10
);
// 定时器实例
let timeoutTimer: ReturnType<typeof setTimeout> | null = null;

/**
 * 记录用户最后操作时间（同步到sessionStorage）
 */
export const recordLastOperateTime = () => {
  lastOperateTime = Date.now();
  sessionStorage.setItem("lastOperateTime", `${lastOperateTime}`);
};

/**
 * 检查是否会话超时（基于sessionStorage的操作时间）
 */
export const checkSessionTimeout = (): boolean => {
  const now = Date.now();
  return now - lastOperateTime > SESSION_TIMEOUT;
};

/**
 * 重置超时定时器
 */
export const resetTimeoutTimer = () => {
  if (timeoutTimer) clearTimeout(timeoutTimer);
  timeoutTimer = setTimeout(handleSessionTimeout, SESSION_TIMEOUT);
};

/**
 * 处理会话超时登出逻辑
 */
export const handleSessionTimeout = () => {
  const loginUserStore = useLoginUserStore();
  if (!loginUserStore.loginUser.id) return;

  message.warning("登录状态已失效，请重新登录");
  // 清空登录态（同步清空sessionStorage）
  loginUserStore.clearLoginUser();
  // 清空操作时间
  sessionStorage.removeItem("lastOperateTime");
  // 跳转登录页并携带原路径
  router.push({
    path: "/user/login",
    query: { redirect: encodeURIComponent(router.currentRoute.value.fullPath) },
  });
};

/**
 * 初始化会话监听：无操作超时 + 页面关闭/刷新处理
 */
export const initSessionTimeoutListener = () => {
  // 1. 监听用户操作，重置超时和操作时间
  const events = ["click", "keydown", "scroll", "mousemove"];
  events.forEach((event) => {
    window.addEventListener(event, () => {
      recordLastOperateTime();
      resetTimeoutTimer();
    });
  });

  // 2. 页面关闭/刷新时：清除定时器（无需清空sessionStorage，关闭页面自动清）
  window.addEventListener("beforeunload", () => {
    if (timeoutTimer) clearTimeout(timeoutTimer);
    timeoutTimer = null;
  });
};
