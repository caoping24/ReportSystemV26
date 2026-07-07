<template>
  <div>
    <div style="display: flex; align-items: center; gap: 16px; margin-bottom: 16px;">
      <a-switch
        v-model:checked="filterEnabled"
        checked-children="筛选"
        un-checked-children="原始"
        :loading="switchLoading"
        @change="handleFilterToggle"
        class="filter-toggle-switch"
      />
    <span v-if="loadedAt" style="color: #999; font-size: 13px;">
      上一次更新规则：{{ formatTime(loadedAt) }} | 共 {{ rules.length }} 条规则
      <a-button type="link" size="small" :loading="loading" @click="handleReload" style="padding: 0 4px;">刷新</a-button>
    </span>
    <span v-else style="color: #999; font-size: 13px;">尚未加载筛选配置</span>
    <span style="margin-left: auto;">
          <a-button type="primary" :loading="savingAll" @click="handleSaveAll" style="margin-left: auto;">保存全部</a-button>
    </span>
    </div>

    <a-table
      :dataSource="rules"
      :columns="columns"
      rowKey="fieldName"
      :pagination="{ pageSize: 50 }"
      :loading="loading"
      size="small"
      bordered
    >
      <template #bodyCell="{ column, record }">
        <template v-if="column.key === 'index'">
          {{ record.fieldName?.replace('Cell', '') }}
        </template>
        <template v-if="column.key === 'comment'">
          <a-input v-model:value="record.comment" size="small" style="width: 80%;" placeholder="添加注释" />
        </template>
        <template v-if="column.key === 'minValue'">
          <a-input-number v-model:value="record.minValue" :min="-9999999" :precision="3" :step="0.001" size="small" style="width: 110px;" :placeholder="'不限'" />
        </template>
        <template v-if="column.key === 'maxValue'">
          <a-input-number v-model:value="record.maxValue" :min="-9999998" :precision="3" :step="0.001" size="small" style="width: 110px;" :placeholder="'不限'" />
        </template>
        <template v-if="column.key === 'action'">
          <a-button type="link" size="small" :loading="record._saving" @click="handleSave(record)">保存</a-button>
        </template>
      </template>
    </a-table>
  </div>
</template>

<script lang="ts" setup>
import { ref, onMounted } from "vue";
import { message } from "ant-design-vue";
import { getFilters, updateConfig, getFilterEnabled, setFilterEnabled } from "@/api/filterConfig";

const loading = ref(false);
const rules = ref<any[]>([]);
const loadedAt = ref("");
const originalRules = ref<string>("");

const columns = [
  { title: "序号", key: "index", width: 60 },
  { title: "字段名", dataIndex: "fieldName", key: "fieldName", width: 50 },
  { title: "注释", dataIndex: "comment", key: "comment", width: 300 },
  { title: "最小值", key: "minValue", width: 120 },
  { title: "最大值", key: "maxValue", width: 120 },
  { title: "操作", key: "action", width: 70 },
];
const savingAll = ref(false);

const filterEnabled = ref(false);
const switchLoading = ref(false);

onMounted(async () => {
  await handleLoadConfig();
  await loadFilterEnabled();
});

async function loadFilterEnabled() {
  try {
    const res = await getFilterEnabled();
    filterEnabled.value = res.data.enabled;
  } catch {
    // 默认 true
  }
}

async function handleFilterToggle(checked: boolean) {
  switchLoading.value = true;
  try {
    await setFilterEnabled(checked);
  } catch {
    filterEnabled.value = !checked;
    message.error("切换失败");
  } finally {
    switchLoading.value = false;
  }
}

async function handleSaveAll() {
  if (rules.value.length === 0) return;

  const origin: any[] = JSON.parse(originalRules.value || "[]");
  const changed = rules.value
    .map((r: any, i: number) => ({ r, o: origin[i] }))
    .filter(({ r, o }: any) =>
      !o || r.minValue !== o.minValue || r.maxValue !== o.maxValue || r.comment !== o.comment
    )
    .map(({ r }: any) => ({
      id: r.id,
      minValue: r.minValue ?? null,
      maxValue: r.maxValue ?? null,
      comment: r.comment ?? null,
    }));

  if (changed.length === 0) {
    message.info("没有修改过的数据");
    return;
  }

  savingAll.value = true;
  let success = 0;
  let fail = 0;

  await Promise.all(changed.map(async (item) => {
    try {
      await updateConfig(item);
      success++;
    } catch {
      fail++;
    }
  }));

  savingAll.value = false;

  if (fail === 0) {
    message.success(`已保存 ${success} 条修改`);
  } else {
    message.warning(`${success} 条成功，${fail} 条失败`);
  }

  await handleLoadConfig();
}

async function handleLoadConfig() {
  loading.value = true;
  try {
    const res = await getFilters();
    const data = res.data;
    if (data.loaded) {
      rules.value = data.rules ?? [];
      originalRules.value = JSON.stringify(data.rules);
      loadedAt.value = data.loadedAt;
    } else {
      rules.value = [];
      originalRules.value = "";
      loadedAt.value = "";
    }
  } catch {
    message.error("获取配置失败");
  } finally {
    loading.value = false;
  }
}

async function handleSave(record: any) {
  record._saving = true;
  try {
    const res = await updateConfig({
      id: record.id,
      minValue: record.minValue ?? null,
      maxValue: record.maxValue ?? null,
      comment: record.comment ?? null,
    });
    message.success(`${record.fieldName} 已更新`);
    loadedAt.value = res.data.loadedAt;
  } catch (e: any) {
    message.error(e?.response?.data?.message ?? "保存失败");
  } finally {
    record._saving = false;
  }
}

async function handleReload() {
  loading.value = true;
  try {
    const { refreshFilters } = await import("@/api/filterConfig");
    const res = await refreshFilters();
    message.success(res.data.message);
    await handleLoadConfig();
    await loadFilterEnabled(); 
  } catch (e: any) {
    message.error(e?.response?.data?.message ?? "刷新失败");
  } finally {
    loading.value = false;
  }
}

function formatTime(iso: string) {
  const d = new Date(iso);
  const pad = (n: number) => n.toString().padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}
</script>

<style scoped>
.filter-toggle-switch.ant-switch-checked {
  background-color: #52c41a !important;
}
.filter-toggle-switch:not(.ant-switch-checked) {
  background-color: #ff4d4f !important;
}
</style>
