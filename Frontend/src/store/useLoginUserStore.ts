import { defineStore } from "pinia";
import { ref } from "vue";
import { getCurrentUser } from "@/api/user";

export const useLoginUserStore = defineStore("loginUser", () => {
  // 初始化：优先从sessionStorage恢复（刷新保留，关闭页面清空）
  const initLoginUser = () => {
    const savedUser = sessionStorage.getItem("loginUser");
    if (savedUser) {
      try {
        return JSON.parse(savedUser);
      } catch (e) {
        console.error("sessionStorage用户信息解析失败：", e);
        return { userName: "未登录" };
      }
    }
    return { userName: "未登录" };
  };
  const loginUser = ref<any>(initLoginUser());

  // 从接口获取当前用户信息（同步到sessionStorage）
  async function fetchLoginUser() {
    try {
      const res = await getCurrentUser();
      if (res.data.code === 0 && res.data.data) {
        console.log("接口获取用户信息：", res.data.data);
        loginUser.value = res.data.data;
        sessionStorage.setItem("loginUser", JSON.stringify(res.data.data));
      }
    } catch (error) {
      console.error("获取当前用户信息失败：", error);
    }
  }

  // 设置登录用户（同步到sessionStorage）
  function setLoginUser(newLoginUser: any) {
    loginUser.value = newLoginUser;
    if (newLoginUser.id) {
      sessionStorage.setItem("loginUser", JSON.stringify(newLoginUser));
    } else {
      sessionStorage.removeItem("loginUser");
    }
  }

  // 清除登录状态（清空sessionStorage）
  function clearLoginUser() {
    loginUser.value = { userName: "未登录" };
    sessionStorage.removeItem("loginUser");
  }

  // 恢复登录状态（从sessionStorage恢复）
  function restoreLoginUser() {
    const savedUser = sessionStorage.getItem("loginUser");
    if (savedUser) {
      try {
        loginUser.value = JSON.parse(savedUser);
      } catch (e) {
        console.error("恢复登录状态失败：", e);
        clearLoginUser();
      }
    }
  }

  return {
    loginUser,
    setLoginUser,
    fetchLoginUser,
    clearLoginUser,
    restoreLoginUser,
  };
});
