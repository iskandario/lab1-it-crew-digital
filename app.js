// Учебный MVP: форма создаёт рекламную задачу и сохраняет её локально.
const taskGrid = document.querySelector('#taskGrid');
const dialog = document.querySelector('#briefDialog');
const form = document.querySelector('#briefForm');
const toast = document.querySelector('#toast');
const storageKey = 'it-crew-tasks-v1';

const defaultTasks = [
  { title: 'Запуск beauty-платформы', goal: 'Трафик', channel: 'Telegram · инфлюенсеры', budget: '180–260 тыс.', offers: 12 },
  { title: '500 регистраций в SaaS', goal: 'Регистрации', channel: 'Telegram Ads · аналитика', budget: '120–190 тыс.', offers: 8 },
  { title: 'Продвижение fashion drop', goal: 'Продажи', channel: 'Блогеры · нативные интеграции', budget: '80–140 тыс.', offers: 15 },
];

function getTasks() { return JSON.parse(localStorage.getItem(storageKey) || 'null') || defaultTasks; }
function renderTasks() {
  taskGrid.innerHTML = getTasks().map((task, index) => `<article class="task-card"><div class="task-number"><span>${String(index + 1).padStart(2, '0')}</span><span>↗</span></div><div><h3>${task.title}</h3><p>${task.goal} · ${task.channel}</p></div><div class="task-meta"><span>Бюджет<strong>${task.budget}</strong></span><span>Предложения<strong>${task.offers}</strong></span></div></article>`).join('');
}
function openBrief() { dialog.showModal(); }
function showToast(message) { toast.textContent = message; toast.classList.add('is-visible'); window.setTimeout(() => toast.classList.remove('is-visible'), 2000); }

['openBriefButton', 'heroBriefButton', 'secondBriefButton', 'agencyBriefButton'].forEach((id) => document.querySelector(`#${id}`).addEventListener('click', openBrief));
form.addEventListener('submit', (event) => {
  event.preventDefault();
  const data = new FormData(form);
  const tasks = getTasks();
  tasks.unshift({ title: data.get('title'), goal: data.get('goal'), channel: data.get('channel'), budget: data.get('budget'), offers: 0 });
  localStorage.setItem(storageKey, JSON.stringify(tasks));
  renderTasks(); dialog.close(); form.reset(); showToast('Задача опубликована в MVP');
});
dialog.addEventListener('click', (event) => { if (event.target === dialog) dialog.close(); });
renderTasks();
