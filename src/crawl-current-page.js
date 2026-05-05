import { chromium } from 'playwright';
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from 'docx';
import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';

const DEFAULT_DEBUG_URL = 'http://127.0.0.1:9222';
const debugUrl = process.env.BROWSER_DEBUG_URL || DEFAULT_DEBUG_URL;
const crawlMode = process.env.CRAWL_MODE || 'batch';
const outputDir = path.resolve('output');
const allQuestionsOutputFile = path.join(outputDir, '全部题目汇总.docx');


const normalizeText = (text) =>
  String(text || '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

const uniqueByText = (items) => {
  const seen = new Set();
  return items.filter((item) => {
    const key = item.text.toLowerCase();
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

const sanitizePathPart = (value) => {
  const sanitized = normalizeText(value)
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return sanitized || '未分类';
};

const normalizeUnitName = (value) => {
  const unitMatch = normalizeText(value).match(/^Unit\s*(\d+)/i);
  return unitMatch ? `Unit${unitMatch[1]}` : sanitizePathPart(value);
};

const buildOutputFileName = (metadata) => {
  const parts = [
    normalizeUnitName(metadata.unit),
    sanitizePathPart(metadata.section),
    sanitizePathPart(metadata.activity),
    metadata.task ? sanitizePathPart(metadata.task) : '',
  ].filter(Boolean);

  return `${parts.join('——') || 'questions'}.docx`;
};


const createParagraph = (text, options = {}) => new Paragraph({
  ...options,
  children: [new TextRun({ text })],
});

const buildQuestionDocument = (result) => {
  const children = [];
  const title = [
    normalizeUnitName(result.category.unit),
    result.category.section,
    result.category.activity,
    result.category.task,
  ].filter(Boolean).join(' —— ');

  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text: title, bold: true })],
  }));

  if (result.directions) {
    children.push(createParagraph(result.directions, { spacing: { after: 240 } }));
  }

  result.questions.forEach((question) => {
    children.push(new Paragraph({
      spacing: { before: 180, after: 80 },
      children: [new TextRun({ text: `${question.questionNumber}. ${question.prompt || ''}`.trim(), bold: true })],
    }));

    question.options.forEach((option) => {
      children.push(createParagraph(`${option.label}. ${option.text}`, {
        indent: { left: 360 },
        spacing: { after: 80 },
      }));
    });
  });

  return new Document({
    styles: {
      default: {
        document: {
          run: { font: 'Arial', size: 24 },
        },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children,
    }],
  });
};

const writeQuestionDocx = async (file, result) => {
  const buffer = await Packer.toBuffer(buildQuestionDocument(result));
  await fs.writeFile(file, buffer);
};

const appendQuestionContent = (children, result) => {
  if (result.directions) {
    children.push(createParagraph(result.directions, { spacing: { after: 180 } }));
  }

  result.questions.forEach((question) => {
    children.push(new Paragraph({
      spacing: { before: 160, after: 80 },
      children: [new TextRun({ text: `${question.questionNumber}. ${question.prompt || ''}`.trim(), bold: true })],
    }));

    question.options.forEach((option) => {
      children.push(createParagraph(`${option.label}. ${option.text}`, {
        indent: { left: 360 },
        spacing: { after: 80 },
      }));
    });
  });
};

const buildAllQuestionsDocument = (items) => {
  const children = [new Paragraph({
    heading: HeadingLevel.TITLE,
    children: [new TextRun({ text: 'Further listening 全部题目汇总', bold: true })],
  })];

  let currentUnit = '';
  let currentSection = '';
  let currentActivity = '';

  items.forEach((item) => {
    const result = item.result;
    const unit = normalizeUnitName(result.category.unit);
    const section = result.category.section;
    const activity = result.category.activity;
    const task = result.category.task;

    if (unit !== currentUnit) {
      currentUnit = unit;
      currentSection = '';
      currentActivity = '';
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 360, after: 180 },
        children: [new TextRun({ text: unit, bold: true })],
      }));
    }

    if (section !== currentSection) {
      currentSection = section;
      currentActivity = '';
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 240, after: 120 },
        children: [new TextRun({ text: section, bold: true })],
      }));
    }

    if (activity !== currentActivity) {
      currentActivity = activity;
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 220, after: 100 },
        children: [new TextRun({ text: activity, bold: true })],
      }));
    }

    if (task) {
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_4,
        spacing: { before: 180, after: 80 },
        children: [new TextRun({ text: task, bold: true })],
      }));
    }

    appendQuestionContent(children, result);
  });

  return new Document({
    styles: {
      default: {
        document: {
          run: { font: 'Arial', size: 24 },
        },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children,
    }],
  });
};

const writeAllQuestionsDocx = async (items) => {
  const buffer = await Packer.toBuffer(buildAllQuestionsDocument(items));
  await fs.writeFile(allQuestionsOutputFile, buffer);
};

async function getActivePage(browser) {
  const contexts = browser.contexts();
  for (const context of contexts) {
    const pages = context.pages().filter((page) => !page.isClosed());
    if (pages.length > 0) return pages.at(-1);
  }
  throw new Error('没有找到已打开的页面。请先用远程调试模式启动浏览器并打开题目页面。');
}

async function extractQuestions(page) {
  return page.evaluate(() => {
    const clean = (value) =>
      String(value || '')
        .replace(/\u00a0/g, ' ')
        .replace(/[ \t]+/g, ' ')
        .replace(/\n{3,}/g, '\n\n')
        .trim();

    const cleanLines = (value) =>
      clean(value)
        .split('\n')
        .map((line) => clean(line))
        .filter(Boolean)
        .filter((line) => !/^\d{1,2}:\d{2}$/.test(line))
        .filter((line) => !/^\d+(\.\d+)?x$/i.test(line))
        .filter((line) => !/^(Words\s*&\s*tips|提\s*交)$/i.test(line));

    const isVisible = (element) => {
      const style = window.getComputedStyle(element);
      const rect = element.getBoundingClientRect();
      return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
    };

    const toSelectorHint = (element) => {
      if (element.id) return `#${element.id}`;
      if (element.className) return `.${String(element.className).trim().split(/\s+/).slice(0, 3).join('.')}`;
      return element.tagName.toLowerCase();
    };

    const parseChoiceQuestions = (text, source, selectorHint) => {
      const lines = cleanLines(text);
      const questionStarts = lines
        .map((line, index) => (/^\d+[.．、)]?$/.test(line) ? index : -1))
        .filter((index) => index >= 0);

      return questionStarts.map((startIndex, orderIndex) => {
        const endIndex = questionStarts[orderIndex + 1] ?? lines.length;
        const block = lines.slice(startIndex, endIndex);
        const questionNumber = block[0].replace(/\D/g, '') || String(orderIndex + 1);
        const content = block.slice(1);
        const options = [];
        let promptLines = [];
        let currentOption = null;

        content.forEach((line) => {
          const optionMatch = line.match(/^([A-H])[.．、)]?$/i);
          if (optionMatch) {
            currentOption = { label: optionMatch[1].toUpperCase(), text: '' };
            options.push(currentOption);
            return;
          }

          if (currentOption) {
            currentOption.text = clean(`${currentOption.text} ${line}`);
            return;
          }

          promptLines.push(line);
        });

        const prompt = clean(promptLines.join('\n'));
        const normalizedOptions = options.filter((option) => option.text);
        const textParts = [];
        textParts.push(prompt ? `${questionNumber}. ${prompt}` : `${questionNumber}.`);
        normalizedOptions.forEach((option) => textParts.push(`${option.label}. ${option.text}`));

        return {
          questionNumber,
          prompt,
          options: normalizedOptions,
          text: textParts.join('\n'),
          source,
          selectorHint,
        };
      }).filter((item) => item.options.length > 0 || item.prompt);
    };

    const directionElement = document.querySelector('.layout-direction-container');
    const directions = directionElement && isVisible(directionElement)
      ? clean(directionElement.innerText || directionElement.textContent)
      : '';

    const focusedSelectors = [
      '.layout-reply-container',
      '.question-wrap',
      '.question-common-abs-question-container .question-wrap',
      '[class*=question i] [class*=reply i]',
      '[class*=question i] [class*=answer i]',
    ];

    for (const selector of focusedSelectors) {
      const element = Array.from(document.querySelectorAll(selector)).find(isVisible);
      if (!element) continue;
      const parsed = parseChoiceQuestions(element.innerText || element.textContent, selector, toSelectorHint(element));
      if (parsed.length > 0) return { directions, questions: parsed };
    }

    const fallbackQuestions = [];
    document.querySelectorAll('p, li, div, h1, h2, h3, h4').forEach((element) => {
      if (!(element instanceof HTMLElement) || !isVisible(element)) return;
      if (element.closest('.pc-slider-menu, .ant-menu, nav, aside, header, footer')) return;
      const text = clean(element.innerText || element.textContent);
      if (text.length < 8 || text.length > 800) return;
      if (!/[?？]$/.test(text) && !/^\s*(\d+|[一二三四五六七八九十]+)[\.、．)]\s*/.test(text)) return;
      fallbackQuestions.push({
        questionNumber: String(fallbackQuestions.length + 1),
        prompt: text,
        options: [],
        text,
        source: 'fallback-text-pattern',
        selectorHint: toSelectorHint(element),
      });
    });

    const seen = new Set();
    const questions = fallbackQuestions.filter((item) => {
      const key = item.text.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

    return { directions, questions };
  });
}

async function extractPageMetadata(page) {
  return page.evaluate(() => {
    const clean = (value) => String(value || '').replace(/\s+/g, ' ').trim();
    const currentMicro = document.querySelector('.pc-slider-menu-micro.pc-menu-activity');
    const menuItems = Array.from(document.querySelectorAll('.pc-slider-menu-unit, .pc-slider-menu-section, .pc-slider-menu-micro'));
    const currentIndex = currentMicro ? menuItems.indexOf(currentMicro) : -1;

    const findPreviousText = (className) => {
      if (currentIndex < 0) return '';
      for (let index = currentIndex - 1; index >= 0; index -= 1) {
        const item = menuItems[index];
        if (item.classList.contains(className)) return clean(item.innerText || item.textContent);
      }
      return '';
    };

    const activeHeaderTab = document.querySelector('.pc-header-tab-activity');
    const activeHeaderTask = document.querySelector('.pc-header-task-activity');
    const activeHeader = activeHeaderTab || activeHeaderTask;
    const activity = currentMicro ? clean(currentMicro.innerText || currentMicro.textContent) : clean(activeHeader?.innerText || activeHeader?.textContent);
    const task = activeHeaderTab && activity && clean(activeHeaderTab.innerText || activeHeaderTab.textContent) !== activity
      ? clean(activeHeaderTab.innerText || activeHeaderTab.textContent)
      : '';

    return {
      unit: findPreviousText('pc-slider-menu-unit'),
      section: findPreviousText('pc-slider-menu-section'),
      activity,
      task,
    };
  });
}

async function collectHeaderTabs(page) {
  return page.evaluate(() => {
    const clean = (value) => String(value || '').replace(/\s+/g, ' ').trim();
    const isVisible = (element) => {
      const style = window.getComputedStyle(element);
      const rect = element.getBoundingClientRect();
      return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
    };

    return Array.from(document.querySelectorAll('.pc-header-tabs-container .tab'))
      .map((element, tabIndex) => ({
        tabIndex,
        text: clean(element.innerText || element.textContent),
        active: element.classList.contains('pc-header-tab-activity'),
      }))
      .filter((item, index) => item.text && isVisible(document.querySelectorAll('.pc-header-tabs-container .tab')[index]));
  });
}

async function clickHeaderTab(page, tab) {
  await dismissBlockingModal(page);
  const tabItem = page.locator('.pc-header-tabs-container .tab').nth(tab.tabIndex);
  await tabItem.scrollIntoViewIfNeeded();
  await clickWithModalRetry(page, tabItem);
  await dismissBlockingModal(page);
  await page.waitForFunction((tabIndex) => {
    const tabs = Array.from(document.querySelectorAll('.pc-header-tabs-container .tab'));
    return tabs[tabIndex]?.classList.contains('pc-header-tab-activity');
  }, tab.tabIndex, { timeout: 10000 });
  await page.waitForLoadState('domcontentloaded');
  await page.waitForTimeout(1200);
  await dismissBlockingModal(page);
}

async function clickWithModalRetry(page, locator) {
  try {
    await locator.click({ timeout: 5000 });
  } catch (error) {
    await dismissBlockingModal(page);
    await locator.click({ timeout: 5000 });
  }
}

async function dismissBlockingModal(page) {
  const buttonTexts = ['我知道了', '知道了', '确定', '确认', 'OK', 'Ok'];

  for (let attempt = 0; attempt < 3; attempt += 1) {
    let dismissed = false;

    for (const text of buttonTexts) {
      const button = page.getByText(text, { exact: true }).last();
      try {
        if (await button.isVisible({ timeout: 300 })) {
          await button.click({ timeout: 2000 });
          dismissed = true;
          break;
        }
      } catch {
        // Continue checking other possible modal buttons.
      }
    }

    if (!dismissed) {
      const modalButton = page.locator('.ant-modal button, .ant-modal-confirm-btns button, [class*=modal i] button').last();
      try {
        if (await modalButton.isVisible({ timeout: 300 })) {
          await modalButton.click({ timeout: 2000 });
          dismissed = true;
        }
      } catch {
        // No visible modal button found.
      }
    }

    if (!dismissed) return false;
    await page.waitForTimeout(500);
  }

  return true;
}

async function collectFurtherListeningTargets(page) {
  return page.evaluate(() => {
    const clean = (value) => String(value || '').replace(/\s+/g, ' ').trim();
    const menuItems = Array.from(document.querySelectorAll('.pc-slider-menu-unit, .pc-slider-menu-section, .pc-slider-menu-micro'));
    const targets = [];
    let currentUnit = '';
    let currentSection = '';

    menuItems.forEach((item, index) => {
      const text = clean(item.innerText || item.textContent);
      if (!text) return;

      if (item.classList.contains('pc-slider-menu-unit')) {
        currentUnit = text;
        currentSection = '';
        return;
      }

      if (item.classList.contains('pc-slider-menu-section')) {
        currentSection = text;
        return;
      }

      if (
        item.classList.contains('pc-slider-menu-micro')
        && currentSection === 'Further listening'
        && ['Conversation', 'Passage', 'Lectures'].includes(text)
        && /^Unit\s+\d+/i.test(currentUnit)
      ) {
        targets.push({
          menuIndex: index,
          unit: currentUnit,
          section: currentSection,
          activity: text,
        });
      }
    });

    return targets;
  });
}

async function clickMenuTarget(page, target) {
  await dismissBlockingModal(page);
  const menuItem = page.locator('.pc-slider-menu-unit, .pc-slider-menu-section, .pc-slider-menu-micro').nth(target.menuIndex);
  await menuItem.scrollIntoViewIfNeeded();
  await clickWithModalRetry(page, menuItem);
  await dismissBlockingModal(page);
  await page.waitForFunction((menuIndex) => {
    const items = Array.from(document.querySelectorAll('.pc-slider-menu-unit, .pc-slider-menu-section, .pc-slider-menu-micro'));
    return items[menuIndex]?.classList.contains('pc-menu-activity');
  }, target.menuIndex, { timeout: 10000 });
  await page.waitForLoadState('domcontentloaded');
  await page.waitForTimeout(1200);
  await dismissBlockingModal(page);
}

async function crawlCurrentPage(page, metadataOverride) {
  await dismissBlockingModal(page);
  const url = page.url();
  const title = normalizeText(await page.title());
  const metadata = metadataOverride || await extractPageMetadata(page);
  const extracted = await extractQuestions(page);
  const questions = uniqueByText(extracted.questions);
  const outputFile = path.join(outputDir, buildOutputFileName(metadata));

  const result = {
    crawledAt: new Date().toISOString(),
    page: { title, url },
    category: metadata,
    directions: extracted.directions,
    count: questions.length,
    questions: questions.map((item, index) => ({ index: index + 1, ...item })),
  };

  await writeQuestionDocx(outputFile, result);

  return { outputFile, result };
}

async function main() {
  await fs.mkdir(outputDir, { recursive: true });

  console.log(`正在连接浏览器：${debugUrl}`);
  const browser = await chromium.connectOverCDP(debugUrl);

  try {
    const page = await getActivePage(browser);
    await page.waitForLoadState('domcontentloaded');

    if (crawlMode === 'current') {
      const { outputFile, result } = await crawlCurrentPage(page);
      console.log(`爬取完成：共 ${result.count} 条，已保存到 ${outputFile}`);
      return;
    }

    const targets = await collectFurtherListeningTargets(page);
    if (targets.length === 0) {
      throw new Error('没有在右侧目录中找到 Further listening 下的 Conversation / Passage / Lectures。');
    }

    console.log(`发现 ${targets.length} 个 Further listening 页面，开始批量爬取...`);
    const summary = [];

    for (const [index, target] of targets.entries()) {
      console.log(`[${index + 1}/${targets.length}] ${target.unit} - ${target.section} - ${target.activity}`);
      await clickMenuTarget(page, target);
      const tabs = target.activity === 'Lectures' ? await collectHeaderTabs(page) : [];

      if (tabs.length > 1) {
        for (const tab of tabs) {
          console.log(`  - ${tab.text}`);
          await clickHeaderTab(page, tab);
          const metadata = { ...target, task: tab.text };
          const { outputFile, result } = await crawlCurrentPage(page, metadata);
          summary.push({
            ...metadata,
            count: result.count,
            outputFile,
            result,
          });
          console.log(`    已保存 ${result.count} 条：${outputFile}`);
        }
        continue;
      }

      const { outputFile, result } = await crawlCurrentPage(page, target);
      summary.push({
        ...target,
        count: result.count,
        outputFile,
        result,
      });
      console.log(`  已保存 ${result.count} 条：${outputFile}`);
    }

    await writeAllQuestionsDocx(summary);
    console.log(`批量爬取完成：共 ${summary.length} 个页面。`);
    console.log(`全部题目汇总已保存到 ${allQuestionsOutputFile}`);
  } finally {
    await browser.close();
  }
}

main().catch((error) => {
  console.error('爬取失败：', error.message);
  process.exitCode = 1;
});





