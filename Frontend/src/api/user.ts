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
export const regenerateReports = async (
  params: { type: number; time: string },
  config?: Record<string, any> // 透传Axios配置（进度监听等）
) => {
  return myAxios.request({
    url: "/api/Report/BuildReport",
    method: "POST",
    responseType: "blob",
    data: params,
    timeout:60000,
    ...config, // 合并进度监听等配置
  });
};

// 批量下载ZIP接口（支持透传配置）
export const batchDownloadReportZip = async (
  params: { type: number; timeStr: string },
  config?: Record<string, any> // 透传Axios配置
) => {
  return myAxios.request({
    url: "/api/File/ZipDownloadFile",
    method: "GET",
    params: params,
    responseType: "blob",
    timeout: 120000,
    ...config, // 合并进度监听等配置
  });
};

// 文件下载工具函数（通用）
export const handleFileDownload = (
  response: { data: Blob },
  fileName: string,
  fileType: 'xlsx' | 'zip'
) => {
  const blob = new Blob([response.data], {
    type: fileType === 'xlsx' 
      ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      : 'application/zip'
  });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  window.URL.revokeObjectURL(url);
  document.body.removeChild(link);
};