import myAxios from "@/request";

export const getLineChartOne = async () => {
  return myAxios.request({
    url: "/api/Dashboard/getLineChartOne",
    method: "GET",
  });
};

export const getLineChartTwo = async () => {
  return myAxios.request({
    url: "/api/Dashboard/getLineChartTwo",
    method: "GET",
  });
};
export const getLineChartThree = async () => {
  return myAxios.request({
    url: "/api/Dashboard/getLineChartThree",
    method: "GET",
  });
};

export const getPieChart = async () => {
  return myAxios.request({
    url: "/api/Dashboard/getPieChart",
    method: "GET",
  });
};

export const getCoreChart = async () => {
  return myAxios.request({
    url: "/api/Dashboard/getCoreChart",
    method: "GET",
  });
};
