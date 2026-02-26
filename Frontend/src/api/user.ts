import myAxios from "@/request";

/**
 * 用户注册
 * @param params
 */
export const userRegister = async (params: any) => {
  return myAxios.request({
    url: "api/user/register",
    method: "POST",
    data: params,
  });
};

/**
 * 用户登录
 * @param params
 */
export const userLogin = async (params: any) => {
  return myAxios.request({
    url: "api/user/login",
    method: "POST",
    data: params,
  });
};

/**
 * 用户注销
 * @param params
 */
export const userLogout = async (params: any) => {
  return myAxios.request({
    url: "api/user/logout",
    method: "POST",
    data: params,
  });
};

/**
 * 获取当前用户
 */
export const getCurrentUser = async () => {
  return myAxios.request({
    url: "api/user/current",
    method: "GET",
  });
};

/**
 * 获取用户列表
 * @param userName
 */
export const searchUsers = async (userName: any) => {
  return myAxios.request({
    url: "api/user/search",
    method: "GET",
    params: {
      userName,
    },
  });
};

/**
 * 删除用户
 * @param id
 */
export const deleteUser = async (id: string) => {
  return myAxios.request({
    url: "api/user/delete",
    method: "POST",
    data: id,
    headers: {
      "Content-Type": "application/json",
    },
  });
};
// 分页查询报表接口
export const getReportByPage = async (params: {
  pageIndex: number;
  pageSize: number;
  Type: number;
}) => {
  return myAxios.request({
    url: "/api/ReportRecord/GetReportByPage", // 对应后端接口地址
    method: "GET",
    params: params,
  });
};

// 你的下载接口封装文件 【精准修复版】
export const downloadReport = async (timeStr: string, tabKey: number) => {
  return myAxios.request({
    url: "api/Report/DownloadExcel",
    method: "GET",
    params: { timeStr: timeStr, Type: tabKey },
    responseType: "blob",
    timeout: 60000,
  });
};
// 新增批量下载报表ZIP接口
// 接口定义文件（如 api/user.ts）
export const batchDownloadReportZip = async (params: {
  type: number; // 报表类型 1-日报 2-周报 3-月报 4-年报
  timeStr: string; // 时间字符串（格式：YYYY/YYYY-MM/YYYY-MM-DD）
}) => {
  return myAxios.request({
    url: "/api/File/ZipDownloadFile", // 后端批量下载接口地址
    method: "GET", // 改为 GET 请求
    params: params, // GET 请求参数放在 params 中（会拼接到 URL）
    responseType: "blob", // 仍需保留 blob 处理二进制流
    timeout: 120000, // 保持超时设置
  });
};
// 新增：重新生成报表接口封装
export const regenerateReports = async (params: {
  type: number; // 报表类型
  time: string; // 时间字符串，后端期望格式请与后端约定（这里建议 'YYYY-MM-DD' 或 'YYYY-MM-01'）
}) => {
  return myAxios.request({
    url: "/api/Report/BuildReport",
    method: "POST",
    data: params,
  });
};
