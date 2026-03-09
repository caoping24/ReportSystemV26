import myAxios from "@/request";



export const getLineChartOne = async (type: any) => {
  return myAxios.request({
    url: "/api/Dashboard/getLineChartOne",
    method: "GET",
    params: {
      type,
    },
  });
};


export const getLineChartTwo = async (type: any) => {
  return myAxios.request({
    url: "/api/Dashboard/getLineChartTwo",
    method: "GET",
       params: {
      type,
    },
  });
};
export const getLineChartThree = async (type: any) => {
  return myAxios.request({
    url: "/api/Dashboard/getLineChartThree",
    method: "GET",
      params: {
      type,
    },
  });
};

export const getPieChart = async (type: any) => {
  return myAxios.request({
    url: "/api/Dashboard/getPieChart",
    method: "GET",
   params: {
      type,
    },
  });
};

export const getCoreChart = async (type: any) => {
  return myAxios.request({
    url: "/api/Dashboard/getCoreChart",
    method: "GET",
     params: {
      type,
    },
  });
};
