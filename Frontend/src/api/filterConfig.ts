import myAxios from "@/request";

/** 刷新筛选配置 */
export const refreshFilters = async () => {
  return myAxios.request({
    url: "/api/FilterConfig/refresh-filters",
    method: "POST",
  });
};

/** 查看当前生效的筛选配置 */
export const getFilters = async () => {
  return myAxios.request({
    url: "/api/FilterConfig/filters",
    method: "GET",
  });
};
export const updateConfig = async (data: {
  id: number;
  minValue: number | null;
  maxValue: number | null;
  comment: string | null;
}) => {
  return myAxios.request({
    url: "/api/FilterConfig/update-config",
    method: "POST",
    data,
  });
};
export const getFilterEnabled = async () => {
  return myAxios.request({
    url: "/api/FilterConfig/filter-enabled",
    method: "GET",
  });
};

export const setFilterEnabled = async (enabled: boolean) => {
  return myAxios.request({
    url: "/api/FilterConfig/filter-enabled",
    method: "POST",
    data: { enabled },
  });
};