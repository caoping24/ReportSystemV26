<template>
  <div id="leaderDashboardPage">
    <!-- 产量指标卡片区 -->
    <div class="top">
      <div class="production-cards">
        <!-- 指标1-5：正常展示数据 -->
        <a-card class="production-card" :loading="isLoading" hoverable>
          <a-statistic title="低压蒸汽消耗(t)" :value="productionData.card1" :precision="2" suffix="">
            <template #prefix><CalendarOutlined class="stat-icon" /></template>
          </a-statistic>
        </a-card>
        <a-card class="production-card" :loading="isLoading" hoverable>
          <a-statistic title="中压蒸汽消耗(t)" :value="productionData.card2" :precision="2" suffix="">
            <template #prefix><CalendarOutlined class="stat-icon" /></template>
          </a-statistic>
        </a-card>
        <a-card class="production-card" :loading="isLoading" hoverable>
          <a-statistic title="电量消耗(kwh)" :value="productionData.card3" :precision="2" suffix=" ">
            <template #prefix><CalendarOutlined class="stat-icon" /></template>
          </a-statistic>
        </a-card>
        <a-card class="production-card" :loading="isLoading" hoverable>
          <a-statistic title="废液排放量((m³))" :value="productionData.card4" :precision="2" suffix="">
            <template #prefix><CalendarOutlined class="stat-icon" /></template>
          </a-statistic>
        </a-card>
        <a-card class="production-card" :loading="isLoading" hoverable>
         <div class="empty-card-header">
            <CalendarOutlined class="stat-icon" />
            <span class="empty-card-title"></span>
          </div>
          <div class="empty-card-content"></div>
        </a-card>

        <!-- 指标6：空白卡片，仅保留标题和图标 -->
        <a-card class="production-card" :loading="isLoading" hoverable>
          <div class="empty-card-header">
            <CalendarOutlined class="stat-icon" />
            <span class="empty-card-title"></span>
          </div>
          <div class="empty-card-content"></div>
        </a-card>

        <!-- 指标7：空白卡片，仅保留标题和图标 -->
        <a-card class="production-card" :loading="isLoading" hoverable>
          <div class="empty-card-header">
            <CalendarOutlined class="stat-icon" />
            <span class="empty-card-title"></span>
          </div>
          <div class="empty-card-content"></div>
        </a-card>

        <!-- 指标8：空白卡片，仅保留标题和图标 -->
        <a-card class="production-card" :loading="isLoading" hoverable>
          <div class="empty-card-header">
            <CalendarOutlined class="stat-icon" />
            <span class="empty-card-title"></span>
          </div>
          <div class="empty-card-content"></div>
        </a-card>
      </div>
    </div>

    <!-- 4个折线图，每排两个 -->
    <div class="chart-grid">
      <a-card class="chart-card" :loading="chartLoading.line1" :body-style="{ padding: '5px' }">
        <div style="width: 100%; height: 300px"><div ref="lineChartRef1" class="chart-container"></div></div>
      </a-card>
      <a-card class="chart-card" :loading="chartLoading.line2" :body-style="{ padding: '5px' }">
        <div style="width: 100%; height: 300px"><div ref="lineChartRef2" class="chart-container"></div></div>
      </a-card>
      <a-card class="chart-card" :loading="chartLoading.line3" :body-style="{ padding: '5px' }">
        <div style="width: 100%; height: 300px"><div ref="lineChartRef3" class="chart-container"></div></div>
      </a-card>
      <a-card class="chart-card" :loading="chartLoading.line4" :body-style="{ padding: '5px' }">
        <div style="width: 100%; height: 300px"><div ref="lineChartRef4" class="chart-container"></div></div>
      </a-card>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { ref, reactive, onMounted, onUnmounted, nextTick } from "vue";
import { message } from "ant-design-vue";
import { CalendarOutlined } from "@ant-design/icons-vue";
import * as echarts from "echarts";
import {
  GetPage3CoreChart1,
  GetPage3LineChart1,
  GetPage3LineChart2,
  GetPage3LineChart3,
  GetPage3LineChart4,
} from "@/api/Dashboard";

// 类型定义
interface ProductionData {
  card1: number; card2: number; card3: number; card4: number;
  card5: number; card6: number; card7: number; card8: number;
}
interface LineChartData {
  xAxis: string[];
  series: { name: string; data: number[] }[];
}
interface ChartLoading {
  line1: boolean; line2: boolean; line3: boolean; line4: boolean;
}
interface EChartsAxisValue { min: number; max: number; data: number[]; }

// 状态
const isLoading = ref(false);
const chartLoading = reactive<ChartLoading>({
  line1: false, line2: false, line3: false, line4: false,
});
const productionData = reactive<ProductionData>({
  card1: 0, card2: 0, card3: 0, card4: 0,
  card5: 0, card6: 0, card7: 0, card8: 0,
});

// 图表ref与实例 (Line1~Line4)
const lineChartRef1 = ref<HTMLDivElement | null>(null);
const lineChartRef2 = ref<HTMLDivElement | null>(null);
const lineChartRef3 = ref<HTMLDivElement | null>(null);
const lineChartRef4 = ref<HTMLDivElement | null>(null);
let lineChartInstance1: echarts.ECharts | null = null;
let lineChartInstance2: echarts.ECharts | null = null;
let lineChartInstance3: echarts.ECharts | null = null;
let lineChartInstance4: echarts.ECharts | null = null;

// 图表数据 (Line1~Line4)
const lineChartData1 = ref<LineChartData>({ xAxis: [], series: [] });
const lineChartData2 = ref<LineChartData>({ xAxis: [], series: [] });
const lineChartData3 = ref<LineChartData>({ xAxis: [], series: [] });
const lineChartData4 = ref<LineChartData>({ xAxis: [], series: [] });

// 卡片数据
const GetPage1CoreChart1Data = async () => {
  try {
    const axiosRes = await GetPage3CoreChart1();
    const res = axiosRes.data as { code: number; data: ProductionData; message: string };
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

// Line1 数据
const fetchLineChartData1 = async () => {
  try {
    chartLoading.line1 = true;
    const axiosRes = await GetPage3LineChart1();
    const res = axiosRes.data as { code: number; data: LineChartData; message: string };
    if (res.code === 0 && res.data?.xAxis && res.data?.series) {
      lineChartData1.value = res.data;
    } else {
      throw new Error(res.message || "数据格式异常");
    }
    await nextTick();
    setTimeout(() => initLineChart1(), 100);
  } catch (error) {
    console.error("获取Line1折线图数据失败：", error);
    message.error("昨日羟基乙腈浓度趋势加载失败");
  } finally {
    chartLoading.line1 = false;
  }
};

// Line2 数据
const fetchLineChartData2 = async () => {
  try {
    chartLoading.line2 = true;
    const axiosRes = await GetPage3LineChart2();
    const res = axiosRes.data as { code: number; data: LineChartData; message: string };
    if (res.code === 0 && res.data?.xAxis && res.data?.series) {
      lineChartData2.value = res.data;
    } else {
      throw new Error(res.message || "数据格式异常");
    }
    await nextTick();
    setTimeout(() => initLineChart2(), 100);
  } catch (error) {
    console.error("获取Line2折线图数据失败：", error);
    message.error("本周摩尔比趋势加载失败");
  } finally {
    chartLoading.line2 = false;
  }
};

// Line3 数据
const fetchLineChartData3 = async () => {
  try {
    chartLoading.line3 = true;
    const axiosRes = await GetPage3LineChart3();
    const res = axiosRes.data as { code: number; data: LineChartData; message: string };
    if (res.code === 0 && res.data?.xAxis && res.data?.series) {
      lineChartData3.value = res.data;
    } else {
      throw new Error(res.message || "数据格式异常");
    }
    await nextTick();
    setTimeout(() => initLineChart3(), 100);
  } catch (error) {
    console.error("获取Line3折线图数据失败：", error);
    message.error("本月羟基乙腈配料浓度趋势加载失败");
  } finally {
    chartLoading.line3 = false;
  }
};

// Line4 数据
const fetchLineChartData4 = async () => {
   try {
    chartLoading.line4 = true;
    const axiosRes = await GetPage3LineChart4();
    const res = axiosRes.data as { code: number; data: LineChartData; message: string };
    if (res.code === 0 && res.data?.xAxis && res.data?.series) {
      lineChartData4.value = res.data;
    } else {
      throw new Error(res.message || "数据格式异常");
    }
    await nextTick();
    setTimeout(() => initLineChart4(), 100);
  } catch (error) {
    console.error("获取Line4折线图数据失败：", error);
    message.error("Line4折线图数据加载失败");
  } finally {
    chartLoading.line4 = false;
  }
};

// 图表初始化函数
const getBaseChartOption = (title: string, yAxisName: string, color: string, data: LineChartData) => {
  const xAxisData = data.xAxis.length ? data.xAxis : ["暂无数据"];
  const seriesData = data.series.length ? data.series : [{ name: "数据", data: [0] }];
  return {
    title: { text: title, left: "center", top: 10, textStyle: { fontSize: 16, fontWeight: 600 } },
    color: [color],
    tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
    legend: { orient: "horizontal", top: 40, left: "center" },
    toolbox: { show: true, feature: { saveAsImage: { show: true, title: "下载图片", type: "png" } }, right: 10, top: 10 },
    grid: { left: "3%", right: "4%", bottom: "3%", top: "70px", containLabel: true },
    xAxis: { type: "category", data: xAxisData, axisLine: { lineStyle: { color: "#e8f4fc" } }, axisLabel: { color: "#666" } },
    yAxis: {
      type: "value", name: yAxisName, nameTextStyle: { color: "#003399" },
      axisLine: { lineStyle: { color: "#e8f4fc" } }, axisLabel: { color: "#666" },
      splitLine: { lineStyle: { color: "#e8f4fc" } },
      min: (value: EChartsAxisValue) => Math.floor(value.min),
      max: (value: EChartsAxisValue) => Math.ceil(value.max),
    },
    series: seriesData.map(item => ({
      name: item.name, type: "line", smooth: true, data: item.data,
      showSymbol: title.includes("近7天") ? true : false,
      lineStyle: { width: title.includes("近7天") ? 2 : 1 }
    }))
  };
};

const initLineChart1 = () => {
  if (!lineChartRef1.value) return;
  if (lineChartInstance1) lineChartInstance1.dispose();
  lineChartInstance1 = echarts.init(lineChartRef1.value);
  lineChartInstance1.setOption(getBaseChartOption("", "", "#003399", lineChartData1.value));
};

const initLineChart2 = () => {
  if (!lineChartRef2.value) return;
  if (lineChartInstance2) lineChartInstance2.dispose();
  lineChartInstance2 = echarts.init(lineChartRef2.value);
  lineChartInstance2.setOption(getBaseChartOption("", "-", "#003399", lineChartData2.value));
};

const initLineChart3 = () => {
  if (!lineChartRef3.value) return;
  if (lineChartInstance3) lineChartInstance3.dispose();
  lineChartInstance3 = echarts.init(lineChartRef3.value);
  lineChartInstance3.setOption(getBaseChartOption("", "", "#003399", lineChartData3.value));
};

const initLineChart4 = () => {
  if (!lineChartRef4.value) return;
  if (lineChartInstance4) lineChartInstance4.dispose();
  lineChartInstance4 = echarts.init(lineChartRef4.value);
  lineChartInstance4.setOption(getBaseChartOption("", "", "#003399", lineChartData4.value));
};

// 刷新所有数据
const fetchAllData = async () => {
  isLoading.value = true;
  await Promise.allSettled([
    GetPage1CoreChart1Data(),
    fetchLineChartData1(),
    fetchLineChartData2(),
    fetchLineChartData3(),
    fetchLineChartData4(),
  ]);
  isLoading.value = false;
  message.success("数据刷新请求已发送");
};

// 生命周期
onMounted(async () => {
  await fetchAllData();
  const resizeHandler = () => {
    [lineChartInstance1, lineChartInstance2, lineChartInstance3, lineChartInstance4].forEach(instance => instance?.resize());
  };
  window.addEventListener("resize", resizeHandler);
  onUnmounted(() => {
    window.removeEventListener("resize", resizeHandler);
    [lineChartInstance1, lineChartInstance2, lineChartInstance3, lineChartInstance4].forEach(instance => instance?.dispose());
  });
});
</script>

<style scoped>
#leaderDashboardPage { padding: 16px; background-color: #f5f7fa; min-height: 80vh; }
.top { width: 100%; margin-bottom: 16px; }
.production-cards { display: grid; grid-template-columns: repeat(8, 1fr); gap: 12px; width: 100%; }
.production-card { border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); min-width: 120px; }
.stat-icon { color: #003399; font-size: 20px; }

/* 4个折线图网格布局：每排两个 */
.chart-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; margin-bottom: 16px; }
.chart-card { border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
.chart-container { width: 100%; height: 100%; }

/* 空白卡片样式 */
.empty-card-header {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 12px;
}
.empty-card-title {
  font-size: 14px;
  color: rgba(0, 0, 0, 0.45);
}
.empty-card-content {
  height: 40px;
}

/* 响应式 */
@media (max-width: 1600px) { .production-cards { grid-template-columns: repeat(6, 1fr); } }
@media (max-width: 1200px) { 
  .production-cards { grid-template-columns: repeat(4, 1fr); }
  .chart-grid { grid-template-columns: repeat(2, 1fr); }
}
@media (max-width: 768px) { 
  .production-cards { grid-template-columns: repeat(2, 1fr); }
  .chart-grid { grid-template-columns: 1fr; }
}
@media (max-width: 480px) { .production-cards { grid-template-columns: 1fr; } }
</style>