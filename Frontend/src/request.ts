import axios, { AxiosResponse, AxiosError } from "axios";
import { message } from "ant-design-vue";
import router from "@/router"; // 引入路由实例
import { useLoginUserStore } from "@/store/useLoginUserStore"; // 引入Pinia store
import {
  checkSessionTimeout,
  handleSessionTimeout,
} from "@/utils/sessionTimeout";

// 扩展AxiosResponse类型
declare module "axios" {
  interface AxiosResponse {
    fileBlobData?: Blob;
    fileDownloadName?: string;
  }
}

// 跳转登录页的锁：防止多次401请求导致重复跳转
let isRedirecting = false;

// 创建axios实例

const myAxios = axios.create({
  // 【优化】动态匹配当前页面域名，避免硬编码IP
  baseURL:
    process.env.NODE_ENV === "development"
      ? "http://localhost:5260"
      : window.location.origin, // 生产环境使用当前页面的域名/IP
  timeout: 10000,
  withCredentials: true, // 允许携带cookie，适配session认证
});
// 请求拦截器
myAxios.interceptors.request.use(
  function (config) {
    if (!config.headers["Content-Type"]) {
      config.headers["Content-Type"] = "application/json;charset=UTF-8";
    }
    const loginUserStore = useLoginUserStore();
    // 从内存的loginUser中获取token，而非localStorage
    if (loginUserStore.loginUser.token) {
      config.headers.Authorization = `Bearer ${loginUserStore.loginUser.token}`;
    }
    return config;
  },
  function (error: AxiosError) {
    message.error("请求配置异常，请检查参数或网络设置");
    console.error("请求拦截器错误：", error);
    return Promise.reject(error);
  }
);

// 响应拦截器
myAxios.interceptors.response.use(
  function (response: AxiosResponse): AxiosResponse {
    // 新增：请求成功但会话超时 → 强制登出
    const loginUserStore = useLoginUserStore();
    if (loginUserStore.loginUser.id && checkSessionTimeout()) {
      handleSessionTimeout();
      return response;
    }

    // blob文件下载处理
    if (response.config.responseType === "blob") {
      let fileName = "导出文件.xlsx";
      if (
        response.headers["content-disposition"] ||
        response.headers["Content-Disposition"]
      ) {
        try {
          const disposition =
            response.headers["content-disposition"] ||
            response.headers["Content-Disposition"];
          fileName = decodeURI(disposition.split("filename=")[1]);
          fileName = fileName.replace(/"/g, "");
        } catch (err) {
          fileName = "导出文件.xlsx";
        }
      }
      response.fileBlobData = response.data as Blob;
      response.fileDownloadName = fileName;
      return response;
    }

    // 常规JSON响应处理
    const { data } = response;
    console.log("响应数据：", data);

    // 40100状态码处理（核心修复：添加.value）
    if (data.code === 40100) {
      const isUserCurrentApi =
        response.request.responseURL.includes("user/current");
      const isLoginPage =
        router.currentRoute.value.path.includes("/user/login");

      if (!isUserCurrentApi && !isLoginPage && !isRedirecting) {
        isRedirecting = true;
        message.warning("登录状态已过期，请重新登录");

        const loginUserStore = useLoginUserStore();
        loginUserStore.clearLoginUser(); // 改用clearLoginUser，自动清空sessionStorage

        router
          .push({
            path: "/user/login",
            query: {
              redirect: encodeURIComponent(router.currentRoute.value.fullPath),
            },
          })
          .finally(() => {
            isRedirecting = false;
          });
      }
    }
    // 业务错误处理
    if (data.code && data.code !== 0 && data.code !== 40100) {
      message.error(data.message || `请求失败（错误码：${data.code}）`);
    }

    return response;
  },
  function (error: AxiosError) {
    if (error.response) {
      const status = error.response.status;
      // 401 HTTP状态码处理（核心修复：添加.value）
      if (status === 401 && !isRedirecting) {
        // 修复3：router.currentRoute.path → router.currentRoute.value.path
        const isLoginPage =
          router.currentRoute.value.path.includes("/user/login");
        if (!isLoginPage) {
          isRedirecting = true;
          message.warning("登录状态已过期，请重新登录");

          const loginUserStore = useLoginUserStore();
          loginUserStore.setLoginUser({ id: "" });
          localStorage.removeItem("loginUser");

          // 修复4：router.currentRoute.fullPath → router.currentRoute.value.fullPath
          router
            .push({
              path: "/user/login",
              query: {
                redirect: encodeURIComponent(
                  router.currentRoute.value.fullPath
                ),
              },
            })
            .finally(() => {
              isRedirecting = false;
            });
        }
      } else if (status !== 401) {
        message.error(`请求失败（HTTP状态码：${status}）`);
      }
    } else if (error.request) {
      message.error("网络异常，请检查网络连接");
    } else {
      message.error("请求配置失败：" + error.message);
    }
    console.error("响应拦截器错误：", error);
    return Promise.reject(error);
  }
);

export default myAxios;
