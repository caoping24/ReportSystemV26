<template>
  <div id="leaderDashboardPage">
    <!-- 上部：产量指标卡片区 -->
    <div class="top">
      <div class="production-cards">
        <!-- 原有卡片内容完全不变 -->
        <a-card class="production-card" :loading="isLoading" hoverable>
          <a-statistic
            title="昨日羟基原料浓度"
            :value="productionData.yesterday"
            :precision="2"
            suffix="g/L"
            class="stat-item"
          >
            <template #prefix>
              <CalendarOutlined class="stat-icon" />
            </template>
          </a-statistic>
        </a-card>
        <a-card class="production-card" :loading="isLoading" hoverable>
          <a-statistic
            title="昨日羟基配料浓度"
            :value="productionData.week"
            :precision="2"
            suffix="g/L"
            class="stat-item"
          >
            <template #prefix>
              <CalendarOutlined class="stat-icon" />
            </template>
          </a-statistic>
        </a-card>
        <a-card class="production-card" :loading="isLoading" hoverable>
          <a-statistic
            title="昨日摩尔比"
            :value="productionData.month"
            :precision="2"
            suffix=" "
            class="stat-item"
          >
            <template #prefix>
              <CalendarOutlined class="stat-icon" />
            </template>
          </a-statistic>
        </a-card>
        <a-card class="production-card" :loading="isLoading" hoverable>
          <a-statistic
            title="昨日累计配比"
            :value="productionData.year"
            :precision="2"
            suffix=" "
            class="stat-item"
          >
            <template #prefix>
              <CalendarOutlined class="stat-icon" />
            </template>
          </a-statistic>
        </a-card>
      </div>
      <!-- 移除：子组件原有刷新按钮 -->
      <!-- <div class="refresh-btn-group">
        <a-button type="primary" @click="fetchAllData" :loading="isLoading">刷新</a-button>
      </div> -->
    </div>

    <!-- 中部、下部图表区域：原有内容完全不变 -->
    <div class="chart-section line-charts-section">
      <a-card
        class="chart-card line-chart-card"
        :loading="chartLoading.dayLine"
        :body-style="{ padding: '5px' }"
      >
        <div style="width: 100%; height: 300px">
          <div ref="dayLineChartRef" class="chart-container"></div>
        </div>
      </a-card>
      <a-card
        class="chart-card line-chart-card"
        :loading="chartLoading.weekLine"
        :body-style="{ padding: '5px' }"
      >
        <div style="width: 100%; height: 300px">
          <div ref="weekLineChartRef" class="chart-container"></div>
        </div>
      </a-card>
      <a-card
        class="chart-card line-chart-card"
        :loading="chartLoading.monthLine"
        :body-style="{ padding: '5px' }"
      >
        <div style="width: 100%; height: 300px">
          <div ref="monthLineChartRef" class="chart-container"></div>
        </div>
      </a-card>
    </div>
    <div class="chart-section pie-bar-chart-section">
      <a-card class="chart-card" :loading="chartLoading.pie">
        <div style="width: 100%; height: 300px">
          <div ref="pieChartRef" class="chart-container"></div>
        </div>
      </a-card>
      <a-card class="chart-card" :loading="chartLoading.bar">
        <div style="width: 100%; height: 300px">
          <div ref="barChartRef" class="chart-container"></div>
        </div>
      </a-card>
    </div>
  </div>
</template>

<script lang="ts" setup>
// 原有导入逻辑完全不变（仅新增defineExpose）
import { ref, reactive, onMounted, onUnmounted, nextTick, defineExpose } from "vue";
import { message } from "ant-design-vue";
import { CalendarOutlined } from "@ant-design/icons-vue";
import * as echarts from "echarts";
import {
  getLineChartOne,
  getLineChartTwo,
  getLineChartThree,
  getPieChart,
  getCoreChart,
} from "@/api/Dashboard";
import myAxios from "@/request";

// 原有类型定义完全不变
interface ProductionData {
  yesterday: number;
  week: number;
  month: number;
  year: number;
}
interface PieChartData {
  name: string;
  value: number;
}
interface LineChartData {
  xAxis: string[];
  series: {
    name: string;
    data: number[];
  }[];
}
interface ProductionQueryParams {
  factoryId?: string;
  warehouseId?: string;
  startTime?: string;
  endTime?: string;
}
interface RealApiResponse<T> {
  code: number;
  data: T;
  message: string;
  description: string | null;
}
interface ChartLoading {
  pie: boolean;
  dayLine: boolean;
  weekLine: boolean;
  monthLine: boolean;
  bar: boolean;
}
interface EChartsAxisValue {
  min: number;
  max: number;
  data: number[];
}

// 原有状态管理完全不变
const isLoading = ref<boolean>(false);
const chartLoading = reactive<ChartLoading>({
  pie: false,
  dayLine: false,
  weekLine: false,
  monthLine: false,
  bar: false,
});
const productionData = reactive<ProductionData>({
  yesterday: 0,
  week: 0,
  month: 0,
  year: 0,
});
const pieChartRef = ref<HTMLDivElement | null>(null);
const dayLineChartRef = ref<HTMLDivElement | null>(null);
const weekLineChartRef = ref<HTMLDivElement | null>(null);
const monthLineChartRef = ref<HTMLDivElement | null>(null);
const barChartRef = ref<HTMLDivElement | null>(null);
let pieChartInstance: echarts.ECharts | null = null;
let dayLineChartInstance: echarts.ECharts | null = null;
let weekLineChartInstance: echarts.ECharts | null = null;
let monthLineChartInstance: echarts.ECharts | null = null;
let barChartInstance: echarts.ECharts | null = null;
const pieChartData = ref<PieChartData[]>([]);
const dayLineChartData = ref<LineChartData>({ xAxis: [], series: [] });
const weekLineChartData = ref<LineChartData>({ xAxis: [], series: [] });
const monthLineChartData = ref<LineChartData>({ xAxis: [], series: [] });
const barChartData = ref<LineChartData>({ xAxis: [], series: [] });

// 原有数据请求方法完全不变（一行不改）
const fetchProductionData = async (params?: ProductionQueryParams) => {
  try {
    const axiosRes = await getCoreChart();
    const res = axiosRes.data as RealApiResponse<ProductionData>;
    if (res.code === 0) {
      Object.assign(productionData, res.data);
    } else {
      throw new Error(res.message);
    }
  } catch (error) {
    console.error("获取核心产量数据失败：", error);
    message.error("核心产量数据加载失败");
  }
};
const fetchPieChartData = async (params?: ProductionQueryParams) => {
  try {
    chartLoading.pie = true;
    const axiosRes = await getPieChart();
    const res = axiosRes.data as RealApiResponse<PieChartData[]>;
    if (res.code === 0) {
      pieChartData.value = res.data;
      await nextTick();
      setTimeout(() => safeInitChart("pie"), 100);
    } else {
      throw new Error(res.message);
    }
  } catch (error) {
    console.error("获取饼图数据失败：", error);
    message.error("产量占比图表数据加载失败");
    pieChartData.value = [
      { name: "生产线A", value: 350 },
      { name: "生产线B", value: 280 },
      { name: "生产线C", value: 420 },
      { name: "生产线D", value: 180 },
    ];
    await nextTick();
    setTimeout(() => safeInitChart("pie"), 100);
  } finally {
    chartLoading.pie = false;
  }
};
const fetchDayLineChartData = async (params?: ProductionQueryParams) => {
  try {
    chartLoading.dayLine = true;
    const axiosRes = await getLineChartOne();
    const res = axiosRes.data as RealApiResponse<LineChartData>;
    if (res.code === 0) {
      if (res.data && res.data.xAxis && res.data.series) {
        dayLineChartData.value = res.data;
      } else {
        throw new Error("接口返回数据格式异常");
      }
      await nextTick();
      setTimeout(() => safeInitChart("dayLine"), 100);
    } else {
      throw new Error(res.message || "获取昨日时段产量数据失败");
    }
  } catch (error) {
    console.error("获取日折线图数据失败：", error);
    message.error("昨日时段产量趋势图表数据加载失败");
    dayLineChartData.value = {
      xAxis: ["00:00", "04:00", "08:00", "12:00", "16:00", "20:00"],
      series: [{ name: "羟基乙腈", data: [85, 88, 92, 89, 95, 91] }],
    };
    await nextTick();
    setTimeout(() => safeInitChart("dayLine"), 100);
  } finally {
    chartLoading.dayLine = false;
  }
};
const fetchWeekLineChartData = async (params?: ProductionQueryParams) => {
  try {
    chartLoading.weekLine = true;
    const axiosRes = await getLineChartTwo();
    const res = axiosRes.data as RealApiResponse<LineChartData>;
    if (res.code === 0) {
      if (res.data && res.data.xAxis && res.data.series) {
        weekLineChartData.value = res.data;
      } else {
        throw new Error("接口返回数据格式异常");
      }
      await nextTick();
      setTimeout(() => safeInitChart("weekLine"), 100);
    } else {
      throw new Error(res.message || "获取周产量数据失败");
    }
  } catch (error) {
    console.error("获取周折线图数据失败：", error);
    message.error("周产量趋势图表数据加载失败");
    weekLineChartData.value = {
      xAxis: ["周一", "周二", "周三", "周四", "周五", "周六", "周日"],
      series: [
        { name: "摩尔比", data: [1.2, 1.3, 1.1, 1.4, 1.25, 1.35, 1.28] },
      ],
    };
    await nextTick();
    setTimeout(() => safeInitChart("weekLine"), 100);
  } finally {
    chartLoading.weekLine = false;
  }
};
const fetchMonthLineChartData = async (params?: ProductionQueryParams) => {
  try {
    chartLoading.monthLine = true;
    const axiosRes = await getLineChartThree();
    const res = axiosRes.data as RealApiResponse<LineChartData>;
    if (res.code === 0) {
      if (res.data && res.data.xAxis && res.data.series) {
        monthLineChartData.value = res.data;
      } else {
        throw new Error("接口返回数据格式异常");
      }
      await nextTick();
      setTimeout(() => safeInitChart("monthLine"), 100);
    } else {
      throw new Error(res.message || "获取月产量数据失败");
    }
  } catch (error) {
    console.error("获取月折线图数据失败：", error);
    message.error("月产量趋势图表数据加载失败");
    monthLineChartData.value = {
      xAxis: ["1日", "5日", "10日", "15日", "20日", "25日", "30日"],
      series: [{ name: "羟基乙腈配料", data: [82, 85, 88, 86, 90, 89, 91] }],
    };
    await nextTick();
    setTimeout(() => safeInitChart("monthLine"), 100);
  } finally {
    chartLoading.monthLine = false;
  }
};
const fetchBarChartData = async (params?: ProductionQueryParams) => {
  try {
    chartLoading.bar = true;
    const mockBarData: LineChartData = {
      xAxis: ["批次1", "批次2", "批次3", "批次4", "批次5", "批次6"],
      series: [
        { name: "羟基原料浓度", data: [85.2, 88.7, 90.1, 87.5, 92.3, 89.8] },
        { name: "羟基配料浓度", data: [78.5, 81.2, 83.7, 80.9, 85.1, 82.4] },
      ],
    };
    barChartData.value = mockBarData;
    await nextTick();
    setTimeout(() => safeInitChart("bar"), 100);
  } catch (error) {
    console.error("获取柱状图数据失败：", error);
    message.error("柱状图数据加载失败");
    barChartData.value = {
      xAxis: ["暂无数据"],
      series: [{ name: "产量", data: [0] }],
    };
    await nextTick();
    setTimeout(() => safeInitChart("bar"), 100);
  } finally {
    chartLoading.bar = false;
  }
};

// 原有图表初始化方法完全不变
const safeInitChart = (
  chartType: "pie" | "dayLine" | "weekLine" | "monthLine" | "bar" | "all" = "all"
) => {
  if (chartType === "pie" || chartType === "all") {
    if (!pieChartRef.value) return;
    try {
      if (pieChartInstance) pieChartInstance.dispose();
      pieChartInstance = echarts.init(pieChartRef.value);
      const pieData = pieChartData.value.length ? pieChartData.value : [{ name: "暂无数据", value: 1 }];
      pieChartInstance.setOption({
        title: { text: "产量占比", left: "center", top: 10, textStyle: { fontSize: 16, fontWeight: 600 } },
        color: ["#003399", "#00AEEF", "#0066CC", "#66B2FF"],
        tooltip: { trigger: "item", formatter: "{a} <br/>{b}: {c} 件 ({d}%)" },
        legend: { orient: "horizontal", bottom: 0, textStyle: { color: "#333" } },
        toolbox: {
          show: true,
          feature: { saveAsImage: { show: true, title: "下载图片", type: "png", pixelRatio: 2, backgroundColor: "#ffffff" } },
          right: 10, top: 10
        },
        series: [{
          name: "产量占比", type: "pie", radius: ["40%", "70%"],
          avoidLabelOverlap: false, label: { show: false },
          emphasis: { label: { show: true, fontSize: 16, fontWeight: 600 } },
          labelLine: { show: false }, data: pieData
        }]
      });
    } catch (error) {
      console.error("初始化饼图失败：", error);
      pieChartInstance = null;
    }
  }
  if (chartType === "dayLine" || chartType === "all") {
    if (!dayLineChartRef.value) return;
    try {
      if (dayLineChartInstance) dayLineChartInstance.dispose();
      dayLineChartInstance = echarts.init(dayLineChartRef.value);
      const xAxisData = dayLineChartData.value.xAxis.length ? dayLineChartData.value.xAxis : ["暂无数据"];
      const seriesData = dayLineChartData.value.series.length ? dayLineChartData.value.series : [{ name: "产量", data: [0] }];
      dayLineChartInstance.setOption({
        title: { text: "昨日羟基乙腈浓度趋势", left: "center", top: 10, textStyle: { fontSize: 16, fontWeight: 600 } },
        color: ["#003399"],
        tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
        legend: { orient: "horizontal", top: 40, left: "center", textStyle: { color: "#333", fontSize: 12 } },
        toolbox: {
          show: true,
          feature: { saveAsImage: { show: true, title: "下载图片", type: "png", pixelRatio: 2, backgroundColor: "#ffffff" } },
          right: 10, top: 10
        },
        grid: { left: "3%", right: "4%", bottom: "3%", top: "70px", containLabel: true },
        xAxis: { type: "category", data: xAxisData, axisLine: { lineStyle: { color: "#e8f4fc" } }, axisLabel: { color: "#666" } },
        yAxis: {
          type: "value", name: "g/L", nameTextStyle: { color: "#003399" },
          axisLine: { lineStyle: { color: "#e8f4fc" } }, axisLabel: { color: "#666" },
          splitLine: { lineStyle: { color: "#e8f4fc" } },
          min: (value: EChartsAxisValue) => Math.floor(value.min),
          max: (value: EChartsAxisValue) => Math.ceil(value.max),
        },
        series: seriesData.map((item) => ({
          name: item.name, type: "line", smooth: true, data: item.data, showSymbol: false, lineStyle: { width: 1 }
        }))
      });
    } catch (error) {
      console.error("初始化日折线图失败：", error);
      dayLineChartInstance = null;
    }
  }
  if (chartType === "weekLine" || chartType === "all") {
    if (!weekLineChartRef.value) return;
    try {
      if (weekLineChartInstance) weekLineChartInstance.dispose();
      weekLineChartInstance = echarts.init(weekLineChartRef.value);
      const xAxisData = weekLineChartData.value.xAxis.length ? weekLineChartData.value.xAxis : ["暂无数据"];
      const seriesData = weekLineChartData.value.series.length ? weekLineChartData.value.series : [{ name: "产量", data: [0] }];
      weekLineChartInstance.setOption({
        title: { text: "本周摩尔比趋势", left: "center", top: 10, textStyle: { fontSize: 16, fontWeight: 600 } },
        color: ["#003399"],
        tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
        legend: { orient: "horizontal", top: 40, left: "center", textStyle: { color: "#333", fontSize: 12 } },
        toolbox: {
          show: true,
          feature: { saveAsImage: { show: true, title: "下载图片", type: "png", pixelRatio: 2, backgroundColor: "#ffffff" } },
          right: 10, top: 10
        },
        grid: { left: "3%", right: "4%", bottom: "3%", top: "70px", containLabel: true },
        xAxis: { type: "category", data: xAxisData, axisLine: { lineStyle: { color: "#e8f4fc" } }, axisLabel: { color: "#666" } },
        yAxis: {
          type: "value", name: "-", nameTextStyle: { color: "#003399" },
          axisLine: { lineStyle: { color: "#e8f4fc" } }, axisLabel: { color: "#666" },
          splitLine: { lineStyle: { color: "#e8f4fc" } },
          min: (value: EChartsAxisValue) => Math.floor(value.min),
          max: (value: EChartsAxisValue) => Math.ceil(value.max),
        },
        series: seriesData.map((item) => ({
          name: item.name, type: "line", smooth: true, data: item.data, showSymbol: false, lineStyle: { width: 1 }
        }))
      });
    } catch (error) {
      console.error("初始化周折线图失败：", error);
      weekLineChartInstance = null;
    }
  }
  if (chartType === "monthLine" || chartType === "all") {
    if (!monthLineChartRef.value) return;
    try {
      if (monthLineChartInstance) monthLineChartInstance.dispose();
      monthLineChartInstance = echarts.init(monthLineChartRef.value);
      const xAxisData = monthLineChartData.value.xAxis.length ? monthLineChartData.value.xAxis : ["暂无数据"];
      const seriesData = monthLineChartData.value.series.length ? monthLineChartData.value.series : [{ name: "产量", data: [0] }];
      monthLineChartInstance.setOption({
        title: { text: "本月羟基乙腈配料浓度趋势", left: "center", top: 10, textStyle: { fontSize: 16, fontWeight: 600 } },
        color: ["#003399"],
        tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
        legend: { orient: "horizontal", top: 40, left: "center", textStyle: { color: "#333", fontSize: 12 } },
        toolbox: {
          show: true,
          feature: { saveAsImage: { show: true, title: "下载图片", type: "png", pixelRatio: 2, backgroundColor: "#ffffff" } },
          right: 10, top: 10
        },
        grid: { left: "3%", right: "4%", bottom: "3%", top: "70px", containLabel: true },
        xAxis: { type: "category", data: xAxisData, axisLine: { lineStyle: { color: "#e8f4fc" } }, axisLabel: { color: "#666" } },
        yAxis: {
          type: "value", name: "g/L", nameTextStyle: { color: "#003399" },
          axisLine: { lineStyle: { color: "#e8f4fc" } }, axisLabel: { color: "#666" },
          splitLine: { lineStyle: { color: "#e8f4fc" } },
          min: (value: EChartsAxisValue) => Math.floor(value.min),
          max: (value: EChartsAxisValue) => Math.ceil(value.max),
        },
        series: seriesData.map((item) => ({
          name: item.name, type: "line", smooth: true, data: item.data, showSymbol: false, lineStyle: { width: 1 }
        }))
      });
    } catch (error) {
      console.error("初始化月折线图失败：", error);
      monthLineChartInstance = null;
    }
  }
  if (chartType === "bar" || chartType === "all") {
    if (!barChartRef.value) return;
    try {
      if (barChartInstance) barChartInstance.dispose();
      barChartInstance = echarts.init(barChartRef.value);
      const xAxisData = barChartData.value.xAxis.length ? barChartData.value.xAxis : ["暂无数据"];
      const seriesData = barChartData.value.series.length ? barChartData.value.series : [{ name: "产量", data: [0] }];
      const colorList = [
        {
          normal: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: "#0066CC" }, { offset: 1, color: "#003399" }]),
          hover: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: "#3399FF" }, { offset: 1, color: "#0066CC" }]),
          border: "#002288",
        },
        {
          normal: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: "#66B2FF" }, { offset: 1, color: "#00AEEF" }]),
          hover: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: "#99CCFF" }, { offset: 1, color: "#66B2FF" }]),
          border: "#0099DD",
        },
      ];
      barChartInstance.setOption({
        title: { text: "各批次浓度对比", left: "center", top: 10, textStyle: { fontSize: 16, fontWeight: 600 } },
        color: [colorList[0].normal, colorList[1].normal],
        tooltip: {
          trigger: "axis", axisPointer: { type: "shadow" },
          textStyle: { fontSize: 12 }, backgroundColor: "rgba(255,255,255,0.9)",
          borderColor: "#e8f4fc", borderWidth: 1, padding: 10,
          formatter: function (params: any) {
            let res = `<div style="font-weight:600;margin-bottom:5px">${params[0].axisValue}</div>`;
            params.forEach((item: any) => {
              res += `<div style="margin:3px 0">
                <span style="display:inline-block;width:8px;height:8px;background:${item.color};margin-right:5px;border-radius:2px;"></span>
                ${item.seriesName}：<span style="font-weight:600">${item.value} g/L</span>
              </div>`;
            });
            return res;
          },
        },
        legend: { orient: "horizontal", top: 40, left: "center", textStyle: { color: "#333", fontSize: 12 }, icon: "rect", itemWidth: 12, itemHeight: 8, itemGap: 20 },
        toolbox: {
          show: true,
          feature: { saveAsImage: { show: true, title: "下载图片", type: "png", pixelRatio: 2, backgroundColor: "#ffffff" } },
          right: 10, top: 10
        },
        grid: { left: "3%", right: "4%", bottom: "3%", top: "70px", containLabel: true },
        xAxis: {
          type: "category", data: xAxisData,
          axisLine: { lineStyle: { color: "#e8f4fc" } },
          axisLabel: { color: "#666", fontSize: 11, interval: 0 },
          axisTick: { alignWithLabel: true }
        },
        yAxis: {
          type: "value", name: "g/L", nameTextStyle: { color: "#0066CC", fontSize: 12 },
          axisLine: { lineStyle: { color: "#e8f4fc" } }, axisLabel: { color: "#666", fontSize: 11 },
          splitLine: { lineStyle: { color: "#e8f4fc", type: "dashed" } },
          min: (value: EChartsAxisValue) => Math.floor(value.min),
          max: (value: EChartsAxisValue) => Math.ceil(value.max),
        },
        series: seriesData.map((item, index) => ({
          name: item.name, type: "bar", barWidth: "35%", barGap: "30%", barCategoryGap: "40%", data: item.data,
          itemStyle: {
            color: colorList[index].normal, borderRadius: [4, 4, 0, 0],
            borderWidth: 1, borderColor: colorList[index].border, borderType: "solid"
          },
          emphasis: {
            itemStyle: {
              color: colorList[index].hover, borderWidth: 1.5,
              shadowBlur: 6, shadowColor: "rgba(0, 102, 204, 0.2)", shadowOffsetY: 2
            },
            label: { show: true, position: "top", fontSize: 11, fontWeight: 600, color: "#333", formatter: "{c} g/L" }
          },
          label: { show: false }
        }))
      });
    } catch (error) {
      console.error("初始化柱状图失败：", error);
      barChartInstance = null;
    }
  }
};

// 原有刷新方法完全不变（仅新增暴露）
const fetchAllData = async () => {
  try {
    isLoading.value = true;
    const params: ProductionQueryParams = {};
    await fetchProductionData(params);
    await Promise.all([
      fetchPieChartData(params),
      fetchDayLineChartData(params),
      fetchWeekLineChartData(params),
      fetchMonthLineChartData(params),
      fetchBarChartData(params),
    ]);
    message.success("数据刷新请求已发送");
    setTimeout(() => {
      safeInitChart("all");
    }, 500);
  } catch (error) {
    console.error("获取核心数据失败：", error);
    message.error("核心数据加载失败，请稍后重试");
  } finally {
    setTimeout(() => {
      isLoading.value = false;
    }, 600);
  }
};

// 原有生命周期完全不变
onMounted(async () => {
  await fetchAllData();
  const resizeHandler = () => {
    if (pieChartInstance) pieChartInstance.resize();
    if (dayLineChartInstance) dayLineChartInstance.resize();
    if (weekLineChartInstance) weekLineChartInstance.resize();
    if (monthLineChartInstance) monthLineChartInstance.resize();
    if (barChartInstance) barChartInstance.resize();
  };
  window.addEventListener("resize", resizeHandler);
  onUnmounted(() => {
    window.removeEventListener("resize", resizeHandler);
    if (pieChartInstance) pieChartInstance.dispose();
    if (dayLineChartInstance) dayLineChartInstance.dispose();
    if (weekLineChartInstance) weekLineChartInstance.dispose();
    if (monthLineChartInstance) monthLineChartInstance.dispose();
    if (barChartInstance) barChartInstance.dispose();
  });
});

// ========== 仅新增：暴露原有fetchAllData方法给父组件 ==========
defineExpose({
  fetchAllData
});
</script>

<style scoped>
/* 原有样式完全不变（仅移除刷新按钮相关样式） */
#leaderDashboardPage {
  padding: 16px;
  background-color: #f5f7fa;
  min-height: 80vh;
}
.top {
  display: flex;
  align-items: flex-start;
  gap: 16px;
  justify-content: space-between;
}
.production-cards {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 16px;
  margin-bottom: 16px;
}
.production-card {
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
}
.stat-item {
  padding: 8px 0;
}
.stat-icon {
  color: #003399;
  font-size: 20px;
}
.chart-section.line-charts-section {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 16px;
  margin-bottom: 16px;
}
.chart-card {
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
}
.line-chart-card {
  height: 300px;
}
.chart-container {
  width: 100%;
  height: 100%;
}
.chart-section.pie-bar-chart-section {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 16px;
  margin-bottom: 10px;
}
/* 移除：原有刷新按钮样式 */
/* .refresh-btn-group {
  width: 120px;
  margin-top: 10px;
  text-align: right;
} */
@media (max-width: 1200px) {
  .production-cards {
    grid-template-columns: repeat(2, 1fr);
  }
  .line-charts-section {
    grid-template-columns: repeat(2, 1fr);
  }
}
@media (max-width: 768px) {
  .production-cards {
    grid-template-columns: 1fr;
  }
  .line-charts-section {
    grid-template-columns: 1fr;
  }
  .pie-bar-chart-section {
    grid-template-columns: 1fr;
  }
  .refresh-btn-group {
    grid-column: 1 / 2;
  }
}
</style>