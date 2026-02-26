<template>
  <div id="userLoginPage">
    <div class="login-card">
      <h2 class="title">用户登录</h2>
      <a-form
        :model="formState"
        label-align="left"
        name="basic"
        :label-col="{ span: 4 }"
        :wrapper-col="{ span: 20 }"
        autocomplete="off"
        @finish="handleSubmit"
        @finishFailed="onFinishFailed"
      >
        <a-form-item
          label="账号"
          name="userAccount"
          :rules="[{ required: true, message: '请输入账号！' }]"
        >
          <a-input
            v-model:value="formState.userAccount"
            placeholder="请输入账号"
          />
        </a-form-item>

        <a-form-item
          label="密码"
          name="userPassword"
          :rules="[
            { required: true, message: '请输入密码!' },
            { min: 8, message: '密码不能小于8位' },
          ]"
        >
          <a-input-password
            v-model:value="formState.userPassword"
            placeholder="请输入密码"
          />
        </a-form-item>

        <a-form-item :wrapper-col="{ offset: 4, span: 20 }">
          <a-button type="primary" html-type="submit" block>登录</a-button>
        </a-form-item>
      </a-form>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { message } from "ant-design-vue";
import { reactive } from "vue";
import { useRouter } from "vue-router";
import { useLoginUserStore } from "@/store/useLoginUserStore";
import { userLogin } from "@/api/user";
import {
  recordLastOperateTime,
  resetTimeoutTimer,
} from "@/utils/sessionTimeout";
const router = useRouter();
const loginUserStore = useLoginUserStore();

// 表单状态
interface FormState {
  userAccount: string;
  userPassword: string;
}
const formState = reactive<FormState>({
  userAccount: "",
  userPassword: "",
});

// 登录提交
const handleSubmit = async (values: any) => {
  try {
    console.log("登录请求参数：", values);
    const res = await userLogin(values);
    console.log("登录接口返回：", res);

    const responseData = res.data || res;
    if (responseData.code === 0 && responseData.data) {
      // 保存用户信息到Store
      loginUserStore.setLoginUser(responseData.data);
      message.success("登录成功");

      // 新增：登录成功后初始化会话超时状态
      recordLastOperateTime();
      resetTimeoutTimer();

      // 延时跳转
      setTimeout(async () => {
        // 优先跳转到redirect参数指定的页面（如果有）
        const redirect = router.currentRoute.value.query.redirect;
        const targetPath = redirect
          ? decodeURIComponent(redirect as string)
          : "/app/components/leader-dashboard";
        await router.push({
          path: targetPath,
          replace: true,
        });
        console.log("登录成功，跳转到目标页面");
      }, 500);
    } else {
      message.error(responseData.message || "登录失败：账号或密码错误");
    }
  } catch (error: any) {
    message.error(error.message || "登录异常，请稍后重试");
    console.error("登录请求失败：", error);
  }
};
// 表单验证失败
const onFinishFailed = (errorInfo: any) => {
  console.log("表单验证失败：", errorInfo);
};
</script>

<style scoped>
#userLoginPage {
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 100vh;
  background: linear-gradient(
      135deg,
      rgba(0, 0, 0, 0.3) 0%,
      rgba(0, 0, 0, 0.3) 100%
    ),
    url("@/assets/login-background.jpg") center/cover no-repeat;
  background-attachment: fixed;
}

.login-card {
  width: 420px;
  padding: 32px;
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.1);
}

.title {
  text-align: center;
  margin-bottom: 24px;
  font-size: 20px;
  font-weight: 600;
  color: #1f2937;
}
</style>
