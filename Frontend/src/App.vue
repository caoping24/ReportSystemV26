<template>
  <div id="app">
    <router-view />
  </div>
</template>

<script setup lang="ts">
// 必须导入onMounted
import { onMounted } from "vue";
import { useLoginUserStore } from "@/store/useLoginUserStore";
import {
  initSessionTimeoutListener,
  checkSessionTimeout,
  handleSessionTimeout,
  recordLastOperateTime,
  resetTimeoutTimer,
} from "@/utils/sessionTimeout";

// 初始化Pinia用户状态（仅内存，无本地恢复）
const loginUserStore = useLoginUserStore();

// 页面挂载后初始化监听
onMounted(() => {
  // 恢复登录状态
  loginUserStore.restoreLoginUser();
  // 仅校验：已登录但超时，直接登出
  if (loginUserStore.loginUser.id && checkSessionTimeout()) {
    handleSessionTimeout();
  }
  recordLastOperateTime();
  initSessionTimeoutListener();
  resetTimeoutTimer();
});
</script>

<style>
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}
html,
body,
#app {
  height: 100%;
}
</style>
