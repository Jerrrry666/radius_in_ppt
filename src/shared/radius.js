/*
 * radius.js — 单位换算常量
 *
 * 锁定信息存到 PPT 文档 customProperty（key = "lock:{shapeId}"）
 * 整个 .pptx 读写通过 Office.js 在 PowerPoint 进程内完成，
 * 所以这个文件只剩共享常量。
 */

const PT_PER_CM = 28.3464567;   // 1 cm = 28.3464567 pt
const ADJ_SCALE = 100000;        // adj value 0~50000 对应 0%~50%

window.RadiusCore = { PT_PER_CM, ADJ_SCALE };
