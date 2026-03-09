import myAxios from "@/request";
//表头
export const Headers = async (params: { type:string }) => {
  return myAxios.request({
    url: "/api/ReportRecord/Headers",
    method: "GET",
    params: params,
  });
};

//列表
export const HourData = async (params: { date: string,type:string }) => {
  return myAxios.request({
    url: "/api/ReportRecord/HourData", // 对应后端接口地址
    method: "GET",
    params: params,
  });
};

export const SaveCell = async (params: any) => {
  return myAxios.request({
    url: "api/ReportRecord/SaveCell",
    method: "POST",
    data: params,
  });
};

export const ReloadData = async (params: any) => {
  return myAxios.request({
    url: "api/Report/BuildReport",
    method: "POST",
    data: params,
  });
};
