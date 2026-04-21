const DEFAULT_ROW_COUNT = 6;
const EPSILON = 1e-9;

const sampleTasks = [
  { id: "A", name: "需求分析", optimistic: 2, mostLikely: 4, pessimistic: 6, predecessors: [] },
  { id: "B", name: "数据库设计", optimistic: 3, mostLikely: 5, pessimistic: 9, predecessors: ["A"] },
  { id: "C", name: "原型设计", optimistic: 2, mostLikely: 3, pessimistic: 5, predecessors: ["A"] },
  { id: "D", name: "后端开发", optimistic: 4, mostLikely: 6, pessimistic: 10, predecessors: ["B"] },
  { id: "E", name: "前端开发", optimistic: 3, mostLikely: 5, pessimistic: 7, predecessors: ["C"] },
  { id: "F", name: "接口联调", optimistic: 2, mostLikely: 4, pessimistic: 6, predecessors: ["D", "E"] },
  { id: "G", name: "测试与修复", optimistic: 3, mostLikely: 4, pessimistic: 8, predecessors: ["F"] },
  { id: "H", name: "部署与演示", optimistic: 1, mostLikely: 2, pessimistic: 3, predecessors: ["F"] },
  { id: "I", name: "文档整理", optimistic: 1, mostLikely: 2, pessimistic: 4, predecessors: ["G", "H"] }
];

const tableBody = document.querySelector("#task-table-body");
const rowTemplate = document.querySelector("#task-row-template");
const messageBox = document.querySelector("#message-box");
const summaryBox = document.querySelector("#summary-box");
const taskResultsBox = document.querySelector("#task-results-box");
const assignmentBox = document.querySelector("#assignment-box");
const ganttBox = document.querySelector("#gantt-box");
const addRowButton = document.querySelector("#add-row-btn");
const loadSampleButton = document.querySelector("#load-sample-btn");
const importButton = document.querySelector("#import-btn");
const exportButton = document.querySelector("#export-btn");
const clearButton = document.querySelector("#clear-btn");
const calculateButton = document.querySelector("#calculate-btn");
const importFileInput = document.querySelector("#import-file-input");

let latestProjectData = null;

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatNumber(value) {
  if (Math.abs(value - Math.round(value)) < EPSILON) {
    return String(Math.round(value));
  }
  return value.toFixed(2).replace(/\.?0+$/, "");
}

function nearlyEqual(a, b) {
  return Math.abs(a - b) < EPSILON;
}

function createRow(task = {}) {
  const rowFragment = rowTemplate.content.cloneNode(true);
  const row = rowFragment.querySelector("tr");

  row.querySelector('input[name="taskId"]').value = task.id ?? "";
  row.querySelector('input[name="taskName"]').value = task.name ?? "";
  row.querySelector('input[name="optimistic"]').value =
    task.optimistic === undefined ? "" : String(task.optimistic);
  row.querySelector('input[name="mostLikely"]').value =
    task.mostLikely === undefined ? "" : String(task.mostLikely);
  row.querySelector('input[name="pessimistic"]').value =
    task.pessimistic === undefined ? "" : String(task.pessimistic);
  row.querySelector('input[name="predecessors"]').value = Array.isArray(task.predecessors)
    ? task.predecessors.join(",")
    : "";

  row.querySelector(".row-remove-btn").addEventListener("click", () => {
    row.remove();
    if (!tableBody.children.length) {
      addEmptyRows(DEFAULT_ROW_COUNT);
    }
  });

  tableBody.appendChild(row);
}

function addEmptyRows(count = 1) {
  for (let index = 0; index < count; index += 1) {
    createRow();
  }
}

function clearTable() {
  tableBody.innerHTML = "";
}

function loadTasks(tasks) {
  clearTable();
  tasks.forEach((task) => createRow(task));
}

function resetOutputs() {
  messageBox.className = "message-box muted";
  messageBox.textContent = "请输入任务数据，然后点击“开始计算”生成 PERT 分析结果。";
  summaryBox.className = "summary-box empty-state";
  summaryBox.textContent = "暂无结果，请先完成一次计算。";
  taskResultsBox.className = "result-box empty-state";
  taskResultsBox.textContent = "这里会显示每个任务的期望工期、时差和关键任务标记。";
  assignmentBox.className = "result-box empty-state";
  assignmentBox.textContent = "这里会显示最少人数条件下的人员排班结果。";
  ganttBox.className = "gantt-box empty-state";
  ganttBox.textContent = "完成计算后，这里会展示按人员分组的时间轴视图。";
  exportButton.disabled = true;
  latestProjectData = null;
}

function showError(message) {
  messageBox.className = "message-box error";
  messageBox.textContent = message;
  summaryBox.className = "summary-box empty-state";
  summaryBox.textContent = "本次计算失败，请修正输入后重试。";
  taskResultsBox.className = "result-box empty-state";
  taskResultsBox.textContent = "没有可展示的任务结果。";
  assignmentBox.className = "result-box empty-state";
  assignmentBox.textContent = "没有可展示的排班结果。";
  ganttBox.className = "gantt-box empty-state";
  ganttBox.textContent = "没有可展示的甘特图。";
  exportButton.disabled = latestProjectData === null;
}

function showSuccess(message) {
  messageBox.className = "message-box success";
  messageBox.textContent = message;
}

function parsePredecessors(rawValue) {
  if (!rawValue.trim()) {
    return [];
  }

  const segments = rawValue.split(",");
  const predecessors = [];

  for (const segment of segments) {
    const trimmed = segment.trim();
    if (!trimmed) {
      throw new Error("前置任务格式有误：请使用英文逗号分隔任务 ID，避免连续逗号。");
    }
    predecessors.push(trimmed);
  }

  return predecessors;
}

function parseNumber(value, fieldLabel, taskId, rowIndex) {
  if (!value.trim()) {
    throw new Error(`第 ${rowIndex + 1} 行任务 ${taskId || "(未命名)"} 缺少 ${fieldLabel}。`);
  }

  const numericValue = Number(value);
  if (!Number.isFinite(numericValue) || numericValue < 0) {
    throw new Error(`第 ${rowIndex + 1} 行任务 ${taskId || "(未命名)"} 的 ${fieldLabel} 必须是非负数字。`);
  }

  return numericValue;
}

function parseTasksFromTable() {
  const rows = Array.from(tableBody.querySelectorAll("tr"));
  const rawTasks = [];

  rows.forEach((row, rowIndex) => {
    const id = row.querySelector('input[name="taskId"]').value.trim();
    const name = row.querySelector('input[name="taskName"]').value.trim();
    const optimisticText = row.querySelector('input[name="optimistic"]').value.trim();
    const mostLikelyText = row.querySelector('input[name="mostLikely"]').value.trim();
    const pessimisticText = row.querySelector('input[name="pessimistic"]').value.trim();
    const predecessorsText = row.querySelector('input[name="predecessors"]').value.trim();

    const isCompletelyBlank =
      !id && !name && !optimisticText && !mostLikelyText && !pessimisticText && !predecessorsText;

    if (isCompletelyBlank) {
      return;
    }

    if (!id) {
      throw new Error(`第 ${rowIndex + 1} 行缺少 Task ID。`);
    }

    if (!name) {
      throw new Error(`第 ${rowIndex + 1} 行任务 ${id} 缺少任务名称。`);
    }

    const optimistic = parseNumber(optimisticText, "O 乐观时间", id, rowIndex);
    const mostLikely = parseNumber(mostLikelyText, "M 最可能时间", id, rowIndex);
    const pessimistic = parseNumber(pessimisticText, "P 悲观时间", id, rowIndex);

    if (optimistic - mostLikely > EPSILON || mostLikely - pessimistic > EPSILON) {
      throw new Error(`任务 ${id} 必须满足 O <= M <= P。`);
    }

    rawTasks.push({
      id,
      name,
      optimistic,
      mostLikely,
      pessimistic,
      predecessors: parsePredecessors(predecessorsText)
    });
  });

  if (!rawTasks.length) {
    throw new Error("请至少输入一条任务数据后再计算。");
  }

  return rawTasks;
}

function buildTaskGraph(rawTasks) {
  const idSet = new Set();
  rawTasks.forEach((task) => {
    if (idSet.has(task.id)) {
      throw new Error(`Task ID 重复：${task.id}。`);
    }
    idSet.add(task.id);
  });

  rawTasks.forEach((task) => {
    task.predecessors.forEach((predecessorId) => {
      if (!idSet.has(predecessorId)) {
        throw new Error(`任务 ${task.id} 引用了不存在的前置任务：${predecessorId}。`);
      }
    });
  });

  const tasks = rawTasks.map((task, inputOrder) => ({
    ...task,
    inputOrder,
    expectedDuration: (task.optimistic + 4 * task.mostLikely + task.pessimistic) / 6,
    successors: []
  }));

  const taskMap = new Map(tasks.map((task) => [task.id, task]));

  tasks.forEach((task) => {
    task.predecessors.forEach((predecessorId) => {
      taskMap.get(predecessorId).successors.push(task.id);
    });
  });

  const indegree = new Map(tasks.map((task) => [task.id, task.predecessors.length]));
  const queue = tasks
    .filter((task) => indegree.get(task.id) === 0)
    .sort((left, right) => left.inputOrder - right.inputOrder)
    .map((task) => task.id);

  const topoOrder = [];

  while (queue.length > 0) {
    const currentId = queue.shift();
    topoOrder.push(currentId);

    taskMap.get(currentId).successors.forEach((successorId) => {
      indegree.set(successorId, indegree.get(successorId) - 1);
      if (indegree.get(successorId) === 0) {
        queue.push(successorId);
        queue.sort((left, right) => taskMap.get(left).inputOrder - taskMap.get(right).inputOrder);
      }
    });
  }

  if (topoOrder.length !== tasks.length) {
    throw new Error("检测到循环依赖，无法执行 PERT 计算。");
  }

  return { tasks, taskMap, topoOrder };
}

function runPertAnalysis(rawTasks) {
  const { tasks, taskMap, topoOrder } = buildTaskGraph(rawTasks);

  topoOrder.forEach((taskId) => {
    const task = taskMap.get(taskId);
    const earliestStart = task.predecessors.length
      ? Math.max(...task.predecessors.map((predecessorId) => taskMap.get(predecessorId).earliestFinish))
      : 0;

    task.earliestStart = earliestStart;
    task.earliestFinish = earliestStart + task.expectedDuration;
  });

  const projectDuration = Math.max(...tasks.map((task) => task.earliestFinish));
  const terminalTasks = tasks.filter((task) => task.successors.length === 0);

  [...topoOrder].reverse().forEach((taskId) => {
    const task = taskMap.get(taskId);
    const latestFinish = task.successors.length
      ? Math.min(...task.successors.map((successorId) => taskMap.get(successorId).latestStart))
      : projectDuration;

    task.latestFinish = latestFinish;
    task.latestStart = latestFinish - task.expectedDuration;
    task.totalFloat = task.latestStart - task.earliestStart;
    task.isCritical = nearlyEqual(task.totalFloat, 0);
  });

  const startTasks = tasks.filter((task) => task.predecessors.length === 0);
  const criticalPaths = [];

  function dfsCriticalPaths(taskId, path) {
    const task = taskMap.get(taskId);
    const nextIds = task.successors.filter((successorId) => {
      const successor = taskMap.get(successorId);
      return successor.isCritical && nearlyEqual(task.earliestFinish, successor.earliestStart);
    });

    if (!nextIds.length) {
      if (nearlyEqual(task.earliestFinish, projectDuration)) {
        criticalPaths.push([...path]);
      }
      return;
    }

    nextIds.forEach((nextId) => {
      dfsCriticalPaths(nextId, [...path, nextId]);
    });
  }

  startTasks
    .filter((task) => task.isCritical && nearlyEqual(task.earliestStart, 0))
    .forEach((task) => dfsCriticalPaths(task.id, [task.id]));

  return {
    tasks: tasks.sort((left, right) => left.inputOrder - right.inputOrder),
    taskMap,
    topoOrder,
    projectDuration,
    criticalPaths,
    terminalTasks: terminalTasks.map((task) => task.id)
  };
}

function getReadyTasks(analysis, scheduledIds, completedIds, currentTime) {
  return analysis.tasks.filter((task) => {
    if (scheduledIds.has(task.id)) {
      return false;
    }
    if (currentTime - task.latestStart > EPSILON) {
      return false;
    }
    if (task.earliestStart - currentTime > EPSILON) {
      return false;
    }
    return task.predecessors.every((predecessorId) => completedIds.has(predecessorId));
  });
}

function getReleasedFutureTasks(analysis, scheduledIds, completedIds, currentTime) {
  return analysis.tasks.filter((task) => {
    if (scheduledIds.has(task.id)) {
      return false;
    }
    if (!task.predecessors.every((predecessorId) => completedIds.has(predecessorId))) {
      return false;
    }
    return task.earliestStart - currentTime > EPSILON;
  });
}

function generateCombinations(items, chooseCount, startIndex = 0, prefix = [], result = []) {
  if (prefix.length === chooseCount) {
    result.push([...prefix]);
    return result;
  }

  for (let index = startIndex; index <= items.length - (chooseCount - prefix.length); index += 1) {
    prefix.push(items[index]);
    generateCombinations(items, chooseCount, index + 1, prefix, result);
    prefix.pop();
  }

  return result;
}

function buildScheduleStateKey(currentTime, completedIds, running) {
  const completedPart = [...completedIds].sort().join(",");
  const runningPart = running
    .map((item) => `${item.taskId}@${formatNumber(item.finishTime)}`)
    .sort()
    .join("|");

  return `${formatNumber(currentTime)}::${completedPart}::${runningPart}`;
}

function searchFeasibleSchedule(analysis, peopleCount) {
  const failedStates = new Set();
  const allTaskIds = new Set(analysis.tasks.map((task) => task.id));

  function recurse(currentTime, completedIds, running, assignments) {
    if (completedIds.size === allTaskIds.size && running.length === 0) {
      return assignments;
    }

    const scheduledIds = new Set(assignments.map((assignment) => assignment.taskId));
    const stateKey = buildScheduleStateKey(currentTime, completedIds, running);
    if (failedStates.has(stateKey)) {
      return null;
    }

    const freePeople = [];
    for (let index = 1; index <= peopleCount; index += 1) {
      const personId = `P${index}`;
      if (!running.some((item) => item.personId === personId)) {
        freePeople.push(personId);
      }
    }

    const readyTasks = getReadyTasks(analysis, scheduledIds, completedIds, currentTime)
      .sort((left, right) => {
        if (!nearlyEqual(left.latestStart, right.latestStart)) {
          return left.latestStart - right.latestStart;
        }
        return left.inputOrder - right.inputOrder;
      });

    const urgentTasks = readyTasks.filter((task) => nearlyEqual(task.latestStart, currentTime));
    if (urgentTasks.length > freePeople.length) {
      failedStates.add(stateKey);
      return null;
    }

    const overdueReadyTask = analysis.tasks.find((task) => {
      if (scheduledIds.has(task.id)) {
        return false;
      }
      if (!task.predecessors.every((predecessorId) => completedIds.has(predecessorId))) {
        return false;
      }
      return currentTime - task.latestStart > EPSILON;
    });

    if (overdueReadyTask) {
      failedStates.add(stateKey);
      return null;
    }

    const optionalReadyTasks = readyTasks.filter(
      (task) => !urgentTasks.some((urgentTask) => urgentTask.id === task.id)
    );
    const maxAdditionalStarts = Math.min(optionalReadyTasks.length, freePeople.length - urgentTasks.length);
    const startSets = [];

    if (freePeople.length > 0 && readyTasks.length > 0) {
      for (let extraCount = maxAdditionalStarts; extraCount >= 0; extraCount -= 1) {
        const combinations = generateCombinations(optionalReadyTasks, extraCount);
        combinations.forEach((combination) => {
          startSets.push([...urgentTasks, ...combination]);
        });
      }
    } else {
      startSets.push([]);
    }

    if (!startSets.length) {
      startSets.push([]);
    }

    for (const startSet of startSets) {
      if (urgentTasks.length > 0 && urgentTasks.some((task) => !startSet.some((item) => item.id === task.id))) {
        continue;
      }

      if (startSet.length === 0 && readyTasks.length > 0 && running.length === 0) {
        continue;
      }

      const orderedStartSet = [...startSet].sort((left, right) => left.inputOrder - right.inputOrder);
      const newAssignments = [...assignments];
      const newRunning = [...running];

      orderedStartSet.forEach((task, index) => {
        const personId = freePeople[index];
        newAssignments.push({
          personId,
          taskId: task.id,
          taskName: task.name,
          startTime: currentTime,
          finishTime: currentTime + task.expectedDuration,
          isCritical: task.isCritical
        });
        newRunning.push({
          personId,
          taskId: task.id,
          finishTime: currentTime + task.expectedDuration
        });
      });

      const newlyScheduledIds = new Set(newAssignments.map((assignment) => assignment.taskId));
      const futureReleaseTimes = getReleasedFutureTasks(
        analysis,
        newlyScheduledIds,
        completedIds,
        currentTime
      ).map((task) => task.earliestStart);

      const eventTimes = [
        ...newRunning.map((item) => item.finishTime),
        ...futureReleaseTimes
      ].filter((time) => time - currentTime > EPSILON);

      if (!eventTimes.length) {
        if (completedIds.size === allTaskIds.size && newRunning.length === 0) {
          return newAssignments;
        }
        continue;
      }

      const nextTime = Math.min(...eventTimes);
      const completedAtNext = new Set(completedIds);
      const remainingRunning = [];

      newRunning.forEach((item) => {
        if (nearlyEqual(item.finishTime, nextTime)) {
          completedAtNext.add(item.taskId);
        } else {
          remainingRunning.push(item);
        }
      });

      const result = recurse(nextTime, completedAtNext, remainingRunning, newAssignments);
      if (result) {
        return result;
      }
    }

    failedStates.add(stateKey);
    return null;
  }

  return recurse(0, new Set(), [], []);
}

function buildResourcePlan(analysis) {
  for (let peopleCount = 1; peopleCount <= analysis.tasks.length; peopleCount += 1) {
    const assignments = searchFeasibleSchedule(analysis, peopleCount);
    if (assignments) {
      const orderedAssignments = assignments.sort((left, right) => {
        if (left.personId !== right.personId) {
          return left.personId.localeCompare(right.personId, undefined, { numeric: true });
        }
        if (!nearlyEqual(left.startTime, right.startTime)) {
          return left.startTime - right.startTime;
        }
        return left.taskId.localeCompare(right.taskId, undefined, { numeric: true });
      });

      return {
        minimumPeople: peopleCount,
        assignments: orderedAssignments
      };
    }
  }

  throw new Error("未找到满足最短工期约束的可行排班方案。");
}

function buildProjectModel(rawTasks) {
  const analysis = runPertAnalysis(rawTasks);
  const resourcePlan = buildResourcePlan(analysis);
  return {
    tasks: analysis.tasks,
    projectDuration: analysis.projectDuration,
    criticalPaths: analysis.criticalPaths,
    minimumPeople: resourcePlan.minimumPeople,
    assignments: resourcePlan.assignments
  };
}

function renderSummary(projectModel) {
  const criticalTaskCount = projectModel.tasks.filter((task) => task.isCritical).length;
  const criticalPathMarkup = projectModel.criticalPaths.length
    ? projectModel.criticalPaths
        .map(
          (path) =>
            `<span class="pill critical">${escapeHtml(path.join(" → "))}</span>`
        )
        .join("")
    : `<span class="pill">未识别到关键路径</span>`;

  summaryBox.className = "summary-box";
  summaryBox.innerHTML = `
    <div class="summary-grid">
      <article class="summary-card">
        <span class="label">项目总工期</span>
        <div class="value">${formatNumber(projectModel.projectDuration)}</div>
      </article>
      <article class="summary-card">
        <span class="label">最少人数</span>
        <div class="value">${projectModel.minimumPeople}</div>
      </article>
      <article class="summary-card">
        <span class="label">关键任务数</span>
        <div class="value">${criticalTaskCount}</div>
      </article>
      <article class="summary-card">
        <span class="label">任务总数</span>
        <div class="value">${projectModel.tasks.length}</div>
      </article>
      <article class="summary-card wide">
        <span class="label">关键路径</span>
        <div class="pill-list">${criticalPathMarkup}</div>
      </article>
    </div>
  `;
}

function renderTaskResults(projectModel) {
  const rows = projectModel.tasks
    .map((task) => {
      const criticalLabel = task.isCritical
        ? `<span class="critical-text">是</span>`
        : "否";

      return `
        <tr class="${task.isCritical ? "critical-row" : ""}">
          <td>${escapeHtml(task.id)}</td>
          <td>${escapeHtml(task.name)}</td>
          <td>${formatNumber(task.optimistic)}</td>
          <td>${formatNumber(task.mostLikely)}</td>
          <td>${formatNumber(task.pessimistic)}</td>
          <td>${formatNumber(task.expectedDuration)}</td>
          <td>${task.predecessors.length ? escapeHtml(task.predecessors.join(", ")) : "-"}</td>
          <td>${formatNumber(task.earliestStart)}</td>
          <td>${formatNumber(task.earliestFinish)}</td>
          <td>${formatNumber(task.latestStart)}</td>
          <td>${formatNumber(task.latestFinish)}</td>
          <td>${formatNumber(task.totalFloat)}</td>
          <td>${criticalLabel}</td>
        </tr>
      `;
    })
    .join("");

  taskResultsBox.className = "result-box";
  taskResultsBox.innerHTML = `
    <p class="helper-line">表中已包含 PERT 期望工期、最早/最晚时间与总时差。关键任务已高亮。</p>
    <div class="table-wrapper">
      <table class="result-table">
        <thead>
          <tr>
            <th>ID</th>
            <th>任务名称</th>
            <th>O</th>
            <th>M</th>
            <th>P</th>
            <th>TE</th>
            <th>前置任务</th>
            <th>ES</th>
            <th>EF</th>
            <th>LS</th>
            <th>LF</th>
            <th>TF</th>
            <th>关键任务</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function renderAssignments(projectModel) {
  const grouped = new Map();
  projectModel.assignments.forEach((assignment) => {
    if (!grouped.has(assignment.personId)) {
      grouped.set(assignment.personId, []);
    }
    grouped.get(assignment.personId).push(assignment);
  });

  const personCards = [...grouped.entries()]
    .map(([personId, assignments]) => {
      const rows = assignments
        .map(
          (assignment) => `
            <tr class="${assignment.isCritical ? "critical-row" : ""}">
              <td>${escapeHtml(assignment.taskId)}</td>
              <td>${escapeHtml(assignment.taskName)}</td>
              <td>${formatNumber(assignment.startTime)}</td>
              <td>${formatNumber(assignment.finishTime)}</td>
              <td>${assignment.isCritical ? `<span class="critical-text">关键任务</span>` : "普通任务"}</td>
            </tr>
          `
        )
        .join("");

      return `
        <section class="assignment-person">
          <h3>${escapeHtml(personId)}</h3>
          <div class="table-wrapper">
            <table class="assignment-table">
              <colgroup>
                <col class="assignment-col-id">
                <col class="assignment-col-name">
                <col class="assignment-col-start">
                <col class="assignment-col-finish">
                <col class="assignment-col-note">
              </colgroup>
              <thead>
                <tr>
                  <th>任务 ID</th>
                  <th>任务名称</th>
                  <th>开始时间</th>
                  <th>结束时间</th>
                  <th>说明</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
        </section>
      `;
    })
    .join("");

  assignmentBox.className = "result-box";
  assignmentBox.innerHTML = `
    <p class="helper-line">
      当前排班满足“项目总工期不变”约束，所需最少人数为 <strong>${projectModel.minimumPeople}</strong> 人。
    </p>
    <div class="assignment-groups">${personCards}</div>
  `;
}

function renderGantt(projectModel) {
  const grouped = new Map();
  projectModel.assignments.forEach((assignment) => {
    if (!grouped.has(assignment.personId)) {
      grouped.set(assignment.personId, []);
    }
    grouped.get(assignment.personId).push(assignment);
  });

  const duration = projectModel.projectDuration || 1;
  const scaleMarkers = [];
  const steps = Math.max(1, Math.min(10, Math.ceil(duration)));

  for (let index = 0; index <= steps; index += 1) {
    const value = (duration / steps) * index;
    scaleMarkers.push(`
      <span class="gantt-scale-marker" style="left:${(value / duration) * 100}%;">
        ${formatNumber(value)}
      </span>
    `);
  }

  const rows = [...grouped.entries()]
    .map(([personId, assignments]) => {
      const bars = assignments
        .map((assignment) => {
          const left = (assignment.startTime / duration) * 100;
          const width = Math.max((assignment.finishTime - assignment.startTime) / duration * 100, 4);
          const isCompact = width < 14;
          const labelLeft = Math.min(left + width + 1, 82);
          const barTitle = `${assignment.taskId} ${assignment.taskName} [${formatNumber(assignment.startTime)} - ${formatNumber(assignment.finishTime)}]`;

          return `
            <div class="gantt-item">
              <div
                class="gantt-bar ${assignment.isCritical ? "critical" : ""} ${isCompact ? "compact" : ""}"
                style="left:${left}%; width:${width}%"
                title="${escapeHtml(barTitle)}"
              >
                <strong>${escapeHtml(assignment.taskId)}</strong>
                <span>${escapeHtml(assignment.taskName)}</span>
              </div>
              ${isCompact ? `
                <div
                  class="gantt-bar-label ${assignment.isCritical ? "critical" : ""}"
                  style="left:${labelLeft}%;"
                  title="${escapeHtml(barTitle)}"
                >
                  ${escapeHtml(`${assignment.taskId} ${assignment.taskName}`)}
                </div>
              ` : ""}
            </div>
          `;
        })
        .join("");

      return `
        <div class="gantt-row">
          <div class="gantt-person-label">${escapeHtml(personId)}</div>
          <div class="gantt-track">${bars}</div>
        </div>
      `;
    })
    .join("");

  ganttBox.className = "gantt-box";
  ganttBox.innerHTML = `
    <div class="gantt-wrapper">
      <div class="gantt-scale">${scaleMarkers.join("")}</div>
      ${rows}
    </div>
  `;
}

function renderProject(projectModel) {
  renderSummary(projectModel);
  renderTaskResults(projectModel);
  renderAssignments(projectModel);
  renderGantt(projectModel);
}

function exportCurrentData() {
  if (!latestProjectData) {
    showError("当前没有可导出的任务数据，请先输入或计算。");
    return;
  }

  const payload = {
    tasks: latestProjectData.tasks.map((task) => ({
      id: task.id,
      name: task.name,
      optimistic: task.optimistic,
      mostLikely: task.mostLikely,
      pessimistic: task.pessimistic,
      predecessors: [...task.predecessors]
    }))
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "pert-project-data.json";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  showSuccess("任务数据已导出为 JSON 文件。");
}

function importTasksFromJsonFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const payload = JSON.parse(reader.result);
      if (!payload || !Array.isArray(payload.tasks)) {
        throw new Error("JSON 文件格式错误：缺少 tasks 数组。");
      }

      const normalizedTasks = payload.tasks.map((task, index) => {
        if (!task || typeof task !== "object") {
          throw new Error(`JSON 第 ${index + 1} 条任务不是有效对象。`);
        }

        return {
          id: task.id ?? "",
          name: task.name ?? "",
          optimistic: task.optimistic,
          mostLikely: task.mostLikely,
          pessimistic: task.pessimistic,
          predecessors: Array.isArray(task.predecessors) ? task.predecessors : []
        };
      });

      loadTasks(normalizedTasks);
      latestProjectData = { tasks: normalizedTasks };
      exportButton.disabled = false;
      showSuccess("JSON 数据已导入，可以直接点击“开始计算”。");
    } catch (error) {
      showError(error.message);
    } finally {
      importFileInput.value = "";
    }
  };

  reader.onerror = () => {
    showError("读取 JSON 文件失败，请重试。");
    importFileInput.value = "";
  };

  reader.readAsText(file, "utf-8");
}

function handleCalculate() {
  try {
    const rawTasks = parseTasksFromTable();
    const projectModel = buildProjectModel(rawTasks);
    latestProjectData = { tasks: rawTasks };
    exportButton.disabled = false;
    renderProject(projectModel);
    showSuccess(
      `计算成功：项目总工期为 ${formatNumber(projectModel.projectDuration)}，最少人数为 ${projectModel.minimumPeople}。`
    );
  } catch (error) {
    showError(error.message);
  }
}

function handleClear() {
  clearTable();
  addEmptyRows(DEFAULT_ROW_COUNT);
  resetOutputs();
}

function initialize() {
  addEmptyRows(DEFAULT_ROW_COUNT);
  exportButton.disabled = true;

  addRowButton.addEventListener("click", () => createRow());
  loadSampleButton.addEventListener("click", () => {
    loadTasks(sampleTasks);
    latestProjectData = { tasks: sampleTasks.map((task) => ({ ...task, predecessors: [...task.predecessors] })) };
    exportButton.disabled = false;
    showSuccess("示例数据已加载，可以直接点击“开始计算”查看完整 PERT 结果。");
  });
  importButton.addEventListener("click", () => importFileInput.click());
  exportButton.addEventListener("click", exportCurrentData);
  clearButton.addEventListener("click", handleClear);
  calculateButton.addEventListener("click", handleCalculate);
  importFileInput.addEventListener("change", (event) => {
    const [file] = event.target.files;
    if (file) {
      importTasksFromJsonFile(file);
    }
  });
}

initialize();
