/**
 * CSS预处理器 - 将<style>标签中的CSS转换为内联样式
 * 解决html-to-docx无法处理外部CSS的问题
 */

import { parseDOM } from 'htmlparser2';
// eslint-disable-next-line import/no-extraneous-dependencies
import render from 'dom-serializer';

/**
 * 解析CSS文本，提取规则
 * @param {string} cssText - CSS文本内容
 * @returns {Object} CSS规则映射 { selector: properties }
 */
function parseCSSRules(cssText) {
  const rules = new Map();

  // 移除注释
  cssText = cssText.replace(/\/\*[\s\S]*?\*\//g, '');

  // 匹配CSS规则: selector { properties }
  // 支持多选择器: h1, h2 { ... }
  const ruleRegex = /([^{]+)\{([^}]+)\}/g;
  let match;

  // eslint-disable-next-line no-cond-assign
  while ((match = ruleRegex.exec(cssText)) !== null) {
    const selectors = match[1].split(',').map((s) => s.trim());
    const properties = match[2].trim();

    selectors.forEach((selector) => {
      // 跳过@规则和伪元素
      if (selector.startsWith('@') || selector.includes('::')) {
        return;
      }

      // 移除伪类
      selector = selector.replace(/:(hover|active|focus|visited|link)/g, '');

      if (selector && properties) {
        if (rules.has(selector)) {
          // 合并相同选择器的规则
          rules.set(selector, `${rules.get(selector)}; ${properties}`);
        } else {
          rules.set(selector, properties);
        }
      }
    });
  }

  return rules;
}

/**
 * 计算CSS选择器的优先级
 * @param {string} selector - CSS选择器
 * @returns {number} 优先级分数
 */
function getSelectorSpecificity(selector) {
  let specificity = 0;

  // ID选择器: 100
  specificity += (selector.match(/#/g) || []).length * 100;

  // 类选择器、属性选择器: 10
  specificity += (selector.match(/\./g) || []).length * 10;
  specificity += (selector.match(/\[/g) || []).length * 10;

  // 标签选择器: 1
  const tagMatches = selector.match(/\b[a-z][a-z0-9]*/gi) || [];
  specificity += tagMatches.length;

  return specificity;
}

/**
 * 检查元素是否匹配选择器
 * @param {Object} element - DOM元素
 * @param {string} selector - CSS选择器
 * @returns {boolean} 是否匹配
 */
function elementMatchesSelector(element, selector) {
  if (!element || !element.type === 'tag') return false;

  selector = selector.trim();

  // 标签选择器
  if (/^[a-z][a-z0-9]*$/i.test(selector)) {
    return element.name === selector.toLowerCase();
  }

  // 类选择器
  if (selector.startsWith('.')) {
    const className = selector.substring(1);
    const classAttr = element.attribs && element.attribs.class;
    if (!classAttr) return false;
    const classes = classAttr.split(/\s+/);
    return classes.includes(className);
  }

  // ID选择器
  if (selector.startsWith('#')) {
    const id = selector.substring(1);
    return element.attribs && element.attribs.id === id;
  }

  // 复合选择器 (简化处理: tag.class)
  const compoundMatch = selector.match(/^([a-z][a-z0-9]*)\.([a-z][a-z0-9-_]*)$/i);
  if (compoundMatch) {
    const [, tag, className] = compoundMatch;
    if (element.name !== tag.toLowerCase()) return false;
    const classAttr = element.attribs && element.attribs.class;
    if (!classAttr) return false;
    const classes = classAttr.split(/\s+/);
    return classes.includes(className);
  }

  return false;
}

/**
 * 为DOM元素应用CSS规则
 * @param {Object} element - DOM元素
 * @param {Map} cssRules - CSS规则映射
 */
function applyStylesToElement(element, cssRules) {
  if (!element || element.type !== 'tag') return;

  // 收集匹配的规则
  const matchedStyles = [];

  cssRules.forEach((properties, selector) => {
    if (elementMatchesSelector(element, selector)) {
      matchedStyles.push({
        selector,
        properties,
        specificity: getSelectorSpecificity(selector),
      });
    }
  });

  // 按优先级排序（低到高）
  matchedStyles.sort((a, b) => a.specificity - b.specificity);

  // 合并样式
  if (matchedStyles.length > 0) {
    const combinedStyles = matchedStyles.map((m) => m.properties).join('; ');

    // 添加到现有style属性
    if (!element.attribs) element.attribs = {};
    const existingStyle = element.attribs.style || '';

    // 现有内联样式优先级最高
    element.attribs.style = existingStyle ? `${combinedStyles}; ${existingStyle}` : combinedStyles;
  }

  // 递归处理子元素
  if (element.children && element.children.length > 0) {
    element.children.forEach((child) => applyStylesToElement(child, cssRules));
  }
}

/**
 * 主函数：预处理HTML，将CSS转换为内联样式
 * @param {string} html - 原始HTML字符串
 * @returns {string} 处理后的HTML字符串
 */
export function preprocessCSS(html) {
  if (!html || typeof html !== 'string') {
    return html;
  }

  // 提取所有<style>标签
  const styleRegex = /<style[^>]*>([\s\S]*?)<\/style>/gi;
  const styleTags = [];
  let match;

  // eslint-disable-next-line no-cond-assign
  while ((match = styleRegex.exec(html)) !== null) {
    styleTags.push(match[1]);
  }

  if (styleTags.length === 0) {
    // 没有样式标签，直接返回
    return html;
  }

  // 合并所有CSS规则
  const allCSS = styleTags.join('\n');
  const cssRules = parseCSSRules(allCSS);

  if (cssRules.size === 0) {
    // 没有有效的CSS规则
    return html;
  }

  // 解析HTML为DOM
  const dom = parseDOM(html);

  // 应用样式到每个元素
  dom.forEach((element) => applyStylesToElement(element, cssRules));

  // 序列化回HTML
  let processedHtml = render(dom, {
    decodeEntities: false,
    encodeEntities: false,
  });

  // 移除<style>标签
  processedHtml = processedHtml.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');

  return processedHtml;
}

/**
 * 简化版本：仅处理常见的简历样式
 * 改进：正确处理多个类和标签选择器，避免重复
 * @param {string} html - 原始HTML
 * @returns {string} 处理后的HTML
 */
export function preprocessResumeCSS(html) {
  let processed = html;

  // 第一步：移除不需要的元素（侧边栏、topbar等）
  // 这些元素在原始HTML中通过display:none或条件类隐藏
  processed = processed.replace(
    /<aside[^>]*class=['"'][^'"]*sidebar[^'"]*['"'][^>]*>[\s\S]*?<\/aside>/gi,
    ''
  );
  processed = processed.replace(
    /<div[^>]*class=['"'][^'"]*topbar[^'"]*['"'][^>]*>[\s\S]*?<\/div>/gi,
    ''
  );

  // 定义简历常用样式映射
  const commonStyles = {
    // 标签选择器
    h1: 'font-size: 22pt; color: #b45309; font-weight: bold; margin: 0;',
    h2: 'font-size: 13pt; color: #b45309; font-weight: bold; margin: 0 0 8pt 0; padding-left: 8pt;',
    h3: 'font-size: 11pt; color: #92400e; font-weight: bold;',
    body: 'font-family: Inter, PingFang SC, Arial, sans-serif; color: #2d1f12;',
    section: 'margin-bottom: 20pt; padding: 16pt; background-color: #fffaf0;',
    strong: 'font-weight: bold;',
    ul: 'margin: 8pt 0;',
    li: 'margin: 4pt 0;',
    div: '', // 空样式，用于后续处理

    // 类选择器 - 注意：为section添加明显的视觉区分
    '.hero': 'margin-bottom: 16pt; padding-bottom: 12pt;',
    '.meta': 'color: #92400e; font-size: 11pt; margin-top: 6pt;',
    '.tag': 'background-color: #fff3cf; color: #92400e; padding: 3pt 7pt; font-size: 10pt;',
    '.title': 'margin: 0;',
    '.sub': 'margin-left: 12pt; color: #666; font-size: 10pt;',
    '.section-block': 'margin-bottom: 20pt; padding: 16pt; background-color: #fffaf0;',
    '.intro': 'background-color: #fffaf0;',
    '.edu': 'background-color: #fffaf0;',
    '.work': 'background-color: #fffaf0;',
    '.proj': 'background-color: #fffaf0;',
    '.skills': 'background-color: #fffaf0;',
    '.cert': 'background-color: #fffaf0;',
    '.other': 'background-color: #fffaf0;',
  };

  // 统一处理：一次性为每个标签应用所有匹配的样式
  processed = processed.replace(/<(\w+)([^>]*?)>/g, (match, tagName, attrs) => {
    const stylesToApply = [];

    // 1. 检查标签选择器
    if (commonStyles[tagName] && commonStyles[tagName] !== '') {
      stylesToApply.push(commonStyles[tagName]);
    }

    // 2. 检查类选择器
    const classMatch = attrs.match(/class=(['"])([^'"]*)\1/);
    if (classMatch) {
      const classList = classMatch[2].split(/\s+/).filter(Boolean);
      classList.forEach((className) => {
        const selector = `.${className}`;
        if (commonStyles[selector]) {
          stylesToApply.push(commonStyles[selector]);
        }
      });
    }

    if (stylesToApply.length === 0) return match; // 没有需要应用的样式

    // 合并所有样式
    const newStyles = stylesToApply.join('; ');

    // 检查是否已有style属性
    const existingStyleMatch = attrs.match(/style=(['"])([^'"]*)\1/);
    if (existingStyleMatch) {
      // 合并到现有style
      const existingStyle = existingStyleMatch[2];
      // 原有内联样式优先级更高，放在后面
      const newAttrs = attrs.replace(
        /style=(['"])([^'"]*)\1/,
        `style=$1${newStyles}; ${existingStyle}$1`
      );
      return `<${tagName}${newAttrs}>`;
    }
    // 添加新style属性
    return `<${tagName}${attrs} style="${newStyles}">`;
  });

  // 移除style标签
  processed = processed.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');

  // 移除main和header等语义标签，保留内容
  processed = processed.replace(/<(main|header|footer|nav|article)([^>]*)>/gi, '<div$2>');
  processed = processed.replace(/<\/(main|header|footer|nav|article)>/gi, '</div>');

  return processed;
}

export default preprocessCSS;
