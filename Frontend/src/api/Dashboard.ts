import myAxios from "@/request";
//2026年5月14日新增获取时间
export function getServerTime() {
  return myAxios.request({
    url: "/api/Dashboard/Now",
    method: "GET",
  });
}
export const GetPage1CoreChart1 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage1CoreChart1",
    method: "GET",
  });
};

export const GetPage1LineChart1 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage1LineChart1",
    method: "GET",
    
  });
};

export const GetPage1LineChart2 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage1LineChart2",
    method: "GET",
    
  });
};
export const GetPage1LineChart3 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage1LineChart3",
    method: "GET",
    
  });
};
export const GetPage1LineChart4 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage1LineChart4",
    method: "GET",
    
  });
};
export const GetPage1LineChart5 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage1LineChart5",
    method: "GET",
    
  });
};
// ------------------第二页接口-------------------------

export const GetPage2CoreChart1 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage2CoreChart1",
    method: "GET",
  });
};

export const GetPage2LineChart1 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage2LineChart1",
    method: "GET",
    
  });
};

export const GetPage2LineChart2 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage2LineChart2",
    method: "GET",
    
  });
};
export const GetPage2LineChart3 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage2LineChart3",
    method: "GET",
    
  });
};
export const GetPage2LineChart4 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage2LineChart4",
    method: "GET",
    
  });
};
export const GetPage2LineChart5 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage2LineChart5",
    method: "GET",
    
  });
};

// -------------第三页

export const GetPage3CoreChart1 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage3CoreChart1",
    method: "GET",
  });
};

export const GetPage3LineChart1 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage3LineChart1",
    method: "GET",
    
  });
};

export const GetPage3LineChart2 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage3LineChart2",
    method: "GET",
    
  });
};
export const GetPage3LineChart3 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage3LineChart3",
    method: "GET",
    
  });
};
export const GetPage3LineChart4 = async () => {
  return myAxios.request({
    url: "/api/Dashboard/GetPage3LineChart4",
    method: "GET",
    
  });
};

