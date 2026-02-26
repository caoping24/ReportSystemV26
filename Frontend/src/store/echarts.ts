// 引入 echarts 核心模块。
import * as echarts from "echarts/core";
// 引入柱状图、折线图、饼图组件。
import { BarChart, LineChart, PieChart } from "echarts/charts";
// 引入标题、提示框、图例、网格、数据集、数据转换器、图形组件。
import {
  TitleComponent,
  TooltipComponent,
  LegendComponent,
  GridComponent,
  DatasetComponent,
  TransformComponent,
  GraphicComponent, // 新增：导入图形组件
} from "echarts/components";
// 引入标签布局和通用过渡动画特性。
import { LabelLayout, UniversalTransition } from "echarts/features";
// 引入 Canvas 渲染器。
import { CanvasRenderer } from "echarts/renderers";

import type {
  // 系列类型的定义后缀都为 SeriesOption
  BarSeriesOption,
  LineSeriesOption,
  PieSeriesOption,
} from "echarts/charts";

import type {
  // 组件类型的定义后缀都为 ComponentOption
  TitleComponentOption,
  TooltipComponentOption,
  LegendComponentOption,
  GridComponentOption,
  DatasetComponentOption,
  GraphicComponentOption, // 新增：图形组件类型
} from "echarts/components";
import type { ComposeOption } from "echarts/core";

// 通过 ComposeOption 来组合出一个只有必须组件和图表的 Option 类型
type ECOption = ComposeOption<
  | BarSeriesOption
  | LineSeriesOption
  | PieSeriesOption
  | TitleComponentOption
  | TooltipComponentOption
  | LegendComponentOption
  | GridComponentOption
  | DatasetComponentOption
  | GraphicComponentOption // 新增：图形组件配置项
>;

/** 
    注册必须的组件，包括标题、提示框、网格、数据集、数据转换器，
    以及柱状图、折线图、饼图、标签布局、通用过渡动画和 Canvas 渲染器。
*/
echarts.use([
  TitleComponent,
  TooltipComponent,
  LegendComponent,
  GridComponent,
  DatasetComponent,
  TransformComponent,
  GraphicComponent, // 新增：注册图形组件
  BarChart,
  LineChart,
  PieChart,
  LabelLayout,
  UniversalTransition,
  CanvasRenderer,
]);
// 导出
export default echarts;
