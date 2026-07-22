/*
 * radius.js — 单位换算常量
 *
 * 锁定信息存到 PPT 文档 customProperty（key = "lock:{shapeId}"）
 * 整个 .pptx 读写通过 Office.js 在 PowerPoint 进程内完成，
 * 所以这个文件只剩共享常量。
 */

const PT_PER_CM = 28.3464567;   // 1 cm = 28.3464567 pt
// Mac LTSC Office.js 在 dialog 上下文里：adjustments.get(0).value 返回 0~1 的小数（占短边比例）
// OOXML 原始是 0~50000（0~50%），但 Office.js 在 Mac dialog 里 normalize 成 0~1 了
const ADJ_SCALE = 1;

window.RadiusCore = { PT_PER_CM, ADJ_SCALE };
