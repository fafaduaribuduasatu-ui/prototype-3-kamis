import { useState, useEffect, useMemo, useRef } from "react";
import { Plus, Radio, Settings2, X, Trash2, Lock, Unlock, ChevronDown, ChevronRight, Crown } from "lucide-react";
import { BarChart, Bar, LineChart, Line, Legend, XAxis, YAxis, ResponsiveContainer, Tooltip, CartesianGrid } from "recharts";
import { FileDown } from "lucide-react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";

const FONT_STYLE = `
@import url('https://fonts.googleapis.com/css2?family=Fredoka:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');
`;

const JAM_OPTIONS = ["05.00","07.00","09.00","11.00","13.00","15.00","17.00","19.00"];
const SESI_OPTIONS = ["SESI 1","SESI 2","SESI 3 (overtime)"];
const TARGET_OPTIONS = ["0 - 2JT","2 - 3JT","3 - 4JT","4 - 5JT","5 - 6JT","7 - 8JT","8 - 9JT","9 - 10JT"];
const DEFAULT_HOSTS = ["ALIN","APRIL","AUFA","HILDA"];

const ACCENT = "#FE2C55";
const CYAN = "#25F4EE";
const GOOD = "#25F4EE";
const BAD = "#FFB800";
const HOST_COLORS = ["#25F4EE", "#FE2C55", "#FFB800", "#8B5CF6", "#4ADE80", "#FF7A5C", "#60A5FA"];
const BASE_UPAH = 25000;
const BONUS_PER_JT = 5000;

function hitungBonus(gmv) {
  const tier = Math.floor((gmv || 0) / 1_000_000);
  return tier >= 2 ? (tier - 1) * BONUS_PER_JT : 0;
}
function hitungTotalGaji(gmv) {
  return BASE_UPAH + hitungBonus(gmv);
}

// Data historis dari Google Form (Juli 2026), diimpor lewat tombol admin
const IMPORT_DATA_JULI_2026 = [
  { host: 'APRIL', tanggal: '2026-07-01', jamMulai: '09.00', sesi: 'SESI 1', gmv: 207000, target: '0 - 2JT', timestamp: '02/07/2026 20:06:57' },
  { host: 'APRIL', tanggal: '2026-07-01', jamMulai: '17.00', sesi: 'SESI 2', gmv: 385000, target: '0 - 2JT', timestamp: '02/07/2026 20:07:44' },
  { host: 'APRIL', tanggal: '2026-07-02', jamMulai: '13.00', sesi: 'SESI 1', gmv: 517000, target: '0 - 2JT', timestamp: '02/07/2026 20:09:18' },
  { host: 'APRIL', tanggal: '2026-07-02', jamMulai: '17.00', sesi: 'SESI 2', gmv: 271000, target: '0 - 2JT', timestamp: '02/07/2026 20:10:05' },
  { host: 'APRIL', tanggal: '2026-07-03', jamMulai: '09.00', sesi: 'SESI 1', gmv: 174000, target: '0 - 2JT', timestamp: '03/07/2026 21:23:03' },
  { host: 'APRIL', tanggal: '2026-07-03', jamMulai: '17.00', sesi: 'SESI 2', gmv: 156000, target: '0 - 2JT', timestamp: '03/07/2026 21:23:31' },
  { host: 'APRIL', tanggal: '2026-07-04', jamMulai: '13.00', sesi: 'SESI 1', gmv: 313000, target: '0 - 2JT', timestamp: '04/07/2026 21:14:36' },
  { host: 'APRIL', tanggal: '2026-07-04', jamMulai: '17.00', sesi: 'SESI 2', gmv: 865000, target: '0 - 2JT', timestamp: '04/07/2026 21:15:09' },
  { host: 'APRIL', tanggal: '2026-07-05', jamMulai: '11.00', sesi: 'SESI 1', gmv: 159000, target: '0 - 2JT', timestamp: '06/07/2026 7:16:16' },
  { host: 'APRIL', tanggal: '2026-07-05', jamMulai: '17.00', sesi: 'SESI 2', gmv: 585000, target: '0 - 2JT', timestamp: '06/07/2026 7:16:50' },
  { host: 'APRIL', tanggal: '2026-07-06', jamMulai: '11.00', sesi: 'SESI 1', gmv: 432000, target: '0 - 2JT', timestamp: '06/07/2026 21:24:35' },
  { host: 'APRIL', tanggal: '2026-07-06', jamMulai: '17.00', sesi: 'SESI 2', gmv: 735000, target: '0 - 2JT', timestamp: '06/07/2026 21:25:21' },
  { host: 'APRIL', tanggal: '2026-07-07', jamMulai: '05.00', sesi: 'SESI 1', gmv: 416000, target: '0 - 2JT', timestamp: '07/07/2026 19:31:02' },
  { host: 'APRIL', tanggal: '2026-07-07', jamMulai: '09.00', sesi: 'SESI 2', gmv: 528000, target: '0 - 2JT', timestamp: '07/07/2026 19:31:32' },
  { host: 'APRIL', tanggal: '2026-07-08', jamMulai: '09.00', sesi: 'SESI 1', gmv: 185000, target: '0 - 2JT', timestamp: '12/07/2026 7:14:09' },
  { host: 'APRIL', tanggal: '2026-07-08', jamMulai: '17.00', sesi: 'SESI 2', gmv: 453000, target: '0 - 2JT', timestamp: '12/07/2026 7:14:57' },
  { host: 'APRIL', tanggal: '2026-07-09', jamMulai: '09.00', sesi: 'SESI 1', gmv: 785000, target: '0 - 2JT', timestamp: '12/07/2026 7:15:40' },
  { host: 'APRIL', tanggal: '2026-07-09', jamMulai: '17.00', sesi: 'SESI 2', gmv: 257000, target: '0 - 2JT', timestamp: '12/07/2026 7:16:25' },
  { host: 'APRIL', tanggal: '2026-07-10', jamMulai: '11.00', sesi: 'SESI 1', gmv: 641000, target: '0 - 2JT', timestamp: '12/07/2026 7:17:13' },
  { host: 'APRIL', tanggal: '2026-07-10', jamMulai: '17.00', sesi: 'SESI 2', gmv: 309000, target: '0 - 2JT', timestamp: '12/07/2026 7:17:52' },
  { host: 'APRIL', tanggal: '2026-07-11', jamMulai: '05.00', sesi: 'SESI 1', gmv: 482000, target: '0 - 2JT', timestamp: '12/07/2026 7:18:29' },
  { host: 'APRIL', tanggal: '2026-07-11', jamMulai: '11.00', sesi: 'SESI 2', gmv: 103000, target: '0 - 2JT', timestamp: '12/07/2026 7:19:09' },
  { host: 'APRIL', tanggal: '2026-07-12', jamMulai: '05.00', sesi: 'SESI 1', gmv: 346000, target: '0 - 2JT', timestamp: '12/07/2026 21:56:59' },
  { host: 'APRIL', tanggal: '2026-07-12', jamMulai: '09.00', sesi: 'SESI 2', gmv: 1200000, target: '0 - 2JT', timestamp: '12/07/2026 21:57:25' },
  { host: 'APRIL', tanggal: '2026-07-13', jamMulai: '15.00', sesi: 'SESI 1', gmv: 143000, target: '0 - 2JT', timestamp: '16/07/2026 20:50:18' },
  { host: 'APRIL', tanggal: '2026-07-13', jamMulai: '19.00', sesi: 'SESI 2', gmv: 360000, target: '0 - 2JT', timestamp: '16/07/2026 20:50:50' },
  { host: 'APRIL', tanggal: '2026-07-14', jamMulai: '05.00', sesi: 'SESI 1', gmv: 478000, target: '0 - 2JT', timestamp: '16/07/2026 20:51:24' },
  { host: 'APRIL', tanggal: '2026-07-14', jamMulai: '11.00', sesi: 'SESI 2', gmv: 100000, target: '0 - 2JT', timestamp: '16/07/2026 20:52:25' },
  { host: 'APRIL', tanggal: '2026-07-15', jamMulai: '05.00', sesi: 'SESI 1', gmv: 299000, target: '0 - 2JT', timestamp: '16/07/2026 20:53:02' },
  { host: 'APRIL', tanggal: '2026-07-15', jamMulai: '11.00', sesi: 'SESI 2', gmv: 421000, target: '0 - 2JT', timestamp: '16/07/2026 20:53:28' },
  { host: 'APRIL', tanggal: '2026-07-16', jamMulai: '05.00', sesi: 'SESI 1', gmv: 200000, target: '0 - 2JT', timestamp: '16/07/2026 20:53:50' },
  { host: 'APRIL', tanggal: '2026-07-16', jamMulai: '11.00', sesi: 'SESI 2', gmv: 241000, target: '0 - 2JT', timestamp: '16/07/2026 20:54:14' },
  { host: 'ALIN', tanggal: '2026-07-01', jamMulai: '07.00', sesi: 'SESI 1', gmv: 265000, target: '0 - 2JT', timestamp: '20/07/2026 9:22:42' },
  { host: 'ALIN', tanggal: '2026-07-01', jamMulai: '13.00', sesi: 'SESI 2', gmv: 408000, target: '0 - 2JT', timestamp: '20/07/2026 9:23:22' },
  { host: 'ALIN', tanggal: '2026-07-02', jamMulai: '07.00', sesi: 'SESI 1', gmv: 301000, target: '0 - 2JT', timestamp: '20/07/2026 9:23:50' },
  { host: 'ALIN', tanggal: '2026-07-02', jamMulai: '13.00', sesi: 'SESI 2', gmv: 630000, target: '0 - 2JT', timestamp: '20/07/2026 9:24:13' },
  { host: 'ALIN', tanggal: '2026-07-03', jamMulai: '07.00', sesi: 'SESI 1', gmv: 278000, target: '0 - 2JT', timestamp: '20/07/2026 9:24:44' },
  { host: 'ALIN', tanggal: '2026-07-03', jamMulai: '13.00', sesi: 'SESI 2', gmv: 301000, target: '0 - 2JT', timestamp: '20/07/2026 9:25:39' },
  { host: 'ALIN', tanggal: '2026-07-04', jamMulai: '05.00', sesi: 'SESI 1', gmv: 219000, target: '0 - 2JT', timestamp: '20/07/2026 9:26:38' },
  { host: 'ALIN', tanggal: '2026-07-04', jamMulai: '09.00', sesi: 'SESI 2', gmv: 399000, target: '0 - 2JT', timestamp: '20/07/2026 9:27:03' },
  { host: 'ALIN', tanggal: '2026-07-05', jamMulai: '07.00', sesi: 'SESI 1', gmv: 527000, target: '0 - 2JT', timestamp: '20/07/2026 9:27:44' },
  { host: 'ALIN', tanggal: '2026-07-05', jamMulai: '13.00', sesi: 'SESI 2', gmv: 443000, target: '0 - 2JT', timestamp: '20/07/2026 9:28:11' },
  { host: 'ALIN', tanggal: '2026-07-06', jamMulai: '07.00', sesi: 'SESI 1', gmv: 179000, target: '0 - 2JT', timestamp: '20/07/2026 9:28:41' },
  { host: 'ALIN', tanggal: '2026-07-06', jamMulai: '13.00', sesi: 'SESI 2', gmv: 501000, target: '0 - 2JT', timestamp: '20/07/2026 9:29:18' },
  { host: 'ALIN', tanggal: '2026-07-07', jamMulai: '07.00', sesi: 'SESI 1', gmv: 695000, target: '0 - 2JT', timestamp: '20/07/2026 9:29:56' },
  { host: 'ALIN', tanggal: '2026-07-07', jamMulai: '13.00', sesi: 'SESI 2', gmv: 410000, target: '0 - 2JT', timestamp: '20/07/2026 9:30:25' },
  { host: 'ALIN', tanggal: '2026-07-08', jamMulai: '07.00', sesi: 'SESI 1', gmv: 878000, target: '0 - 2JT', timestamp: '20/07/2026 9:30:59' },
  { host: 'ALIN', tanggal: '2026-07-08', jamMulai: '13.00', sesi: 'SESI 2', gmv: 536000, target: '0 - 2JT', timestamp: '20/07/2026 9:31:26' },
  { host: 'ALIN', tanggal: '2026-07-09', jamMulai: '07.00', sesi: 'SESI 1', gmv: 794000, target: '0 - 2JT', timestamp: '20/07/2026 9:32:09' },
  { host: 'ALIN', tanggal: '2026-07-09', jamMulai: '13.00', sesi: 'SESI 2', gmv: 749000, target: '0 - 2JT', timestamp: '20/07/2026 9:32:36' },
  { host: 'ALIN', tanggal: '2026-07-10', jamMulai: '07.00', sesi: 'SESI 1', gmv: 262000, target: '0 - 2JT', timestamp: '20/07/2026 9:33:10' },
  { host: 'ALIN', tanggal: '2026-07-10', jamMulai: '13.00', sesi: 'SESI 2', gmv: 755000, target: '0 - 2JT', timestamp: '20/07/2026 9:33:39' },
  { host: 'ALIN', tanggal: '2026-07-11', jamMulai: '07.00', sesi: 'SESI 1', gmv: 640000, target: '0 - 2JT', timestamp: '20/07/2026 9:34:19' },
  { host: 'ALIN', tanggal: '2026-07-11', jamMulai: '13.00', sesi: 'SESI 2', gmv: 751000, target: '0 - 2JT', timestamp: '20/07/2026 9:34:49' },
  { host: 'ALIN', tanggal: '2026-07-12', jamMulai: '07.00', sesi: 'SESI 1', gmv: 600000, target: '0 - 2JT', timestamp: '20/07/2026 9:35:22' },
  { host: 'ALIN', tanggal: '2026-07-12', jamMulai: '13.00', sesi: 'SESI 2', gmv: 1100000, target: '0 - 2JT', timestamp: '20/07/2026 9:35:55' },
  { host: 'ALIN', tanggal: '2026-07-13', jamMulai: '07.00', sesi: 'SESI 1', gmv: 0, target: '0 - 2JT', timestamp: '20/07/2026 9:36:19' },
  { host: 'ALIN', tanggal: '2026-07-13', jamMulai: '13.00', sesi: 'SESI 2', gmv: 460000, target: '0 - 2JT', timestamp: '20/07/2026 9:36:39' },
  { host: 'ALIN', tanggal: '2026-07-14', jamMulai: '07.00', sesi: 'SESI 1', gmv: 208000, target: '0 - 2JT', timestamp: '20/07/2026 9:37:04' },
  { host: 'ALIN', tanggal: '2026-07-14', jamMulai: '13.00', sesi: 'SESI 2', gmv: 559000, target: '0 - 2JT', timestamp: '20/07/2026 9:37:29' },
  { host: 'ALIN', tanggal: '2026-07-15', jamMulai: '07.00', sesi: 'SESI 1', gmv: 103000, target: '0 - 2JT', timestamp: '20/07/2026 9:38:08' },
  { host: 'ALIN', tanggal: '2026-07-15', jamMulai: '13.00', sesi: 'SESI 2', gmv: 301000, target: '0 - 2JT', timestamp: '20/07/2026 9:38:36' },
  { host: 'ALIN', tanggal: '2026-07-16', jamMulai: '09.00', sesi: 'SESI 1', gmv: 719000, target: '0 - 2JT', timestamp: '20/07/2026 9:39:23' },
  { host: 'ALIN', tanggal: '2026-07-16', jamMulai: '13.00', sesi: 'SESI 2', gmv: 308000, target: '0 - 2JT', timestamp: '20/07/2026 9:39:51' },
  { host: 'ALIN', tanggal: '2026-07-17', jamMulai: '09.00', sesi: 'SESI 1', gmv: 591000, target: '0 - 2JT', timestamp: '20/07/2026 9:40:35' },
  { host: 'ALIN', tanggal: '2026-07-17', jamMulai: '13.00', sesi: 'SESI 2', gmv: 243000, target: '0 - 2JT', timestamp: '20/07/2026 9:41:00' },
  { host: 'ALIN', tanggal: '2026-07-18', jamMulai: '07.00', sesi: 'SESI 1', gmv: 364000, target: '0 - 2JT', timestamp: '20/07/2026 9:41:52' },
  { host: 'ALIN', tanggal: '2026-07-19', jamMulai: '07.00', sesi: 'SESI 1', gmv: 295000, target: '0 - 2JT', timestamp: '20/07/2026 9:42:25' },
  { host: 'ALIN', tanggal: '2026-07-19', jamMulai: '13.00', sesi: 'SESI 2', gmv: 656000, target: '0 - 2JT', timestamp: '20/07/2026 9:42:46' },
  { host: 'AUFA', tanggal: '2026-07-01', jamMulai: '11.00', sesi: 'SESI 1', gmv: 632000, target: '0 - 2JT', timestamp: '20/07/2026 11:06:09' },
  { host: 'AUFA', tanggal: '2026-07-02', jamMulai: '15.00', sesi: 'SESI 1', gmv: 397000, target: '0 - 2JT', timestamp: '20/07/2026 11:06:56' },
  { host: 'AUFA', tanggal: '2026-07-02', jamMulai: '19.00', sesi: 'SESI 2', gmv: 511000, target: '0 - 2JT', timestamp: '20/07/2026 11:07:28' },
  { host: 'AUFA', tanggal: '2026-07-03', jamMulai: '15.00', sesi: 'SESI 1', gmv: 607000, target: '0 - 2JT', timestamp: '20/07/2026 11:08:06' },
  { host: 'AUFA', tanggal: '2026-07-03', jamMulai: '19.00', sesi: 'SESI 2', gmv: 258000, target: '0 - 2JT', timestamp: '20/07/2026 11:08:59' },
  { host: 'AUFA', tanggal: '2026-07-04', jamMulai: '15.00', sesi: 'SESI 1', gmv: 517000, target: '0 - 2JT', timestamp: '20/07/2026 11:09:33' },
  { host: 'AUFA', tanggal: '2026-07-04', jamMulai: '19.00', sesi: 'SESI 2', gmv: 774000, target: '0 - 2JT', timestamp: '20/07/2026 11:10:04' },
  { host: 'AUFA', tanggal: '2026-07-05', jamMulai: '09.00', sesi: 'SESI 1', gmv: 622000, target: '0 - 2JT', timestamp: '20/07/2026 11:10:46' },
  { host: 'AUFA', tanggal: '2026-07-06', jamMulai: '15.00', sesi: 'SESI 1', gmv: 629000, target: '0 - 2JT', timestamp: '20/07/2026 11:11:24' },
  { host: 'AUFA', tanggal: '2026-07-06', jamMulai: '09.00', sesi: 'SESI 2', gmv: 593000, target: '0 - 2JT', timestamp: '20/07/2026 11:12:21' },
  { host: 'AUFA', tanggal: '2026-07-07', jamMulai: '15.00', sesi: 'SESI 1', gmv: 614000, target: '0 - 2JT', timestamp: '20/07/2026 11:12:54' },
  { host: 'AUFA', tanggal: '2026-07-07', jamMulai: '19.00', sesi: 'SESI 2', gmv: 440000, target: '0 - 2JT', timestamp: '20/07/2026 11:13:21' },
  { host: 'AUFA', tanggal: '2026-07-08', jamMulai: '15.00', sesi: 'SESI 1', gmv: 766000, target: '0 - 2JT', timestamp: '20/07/2026 11:13:56' },
  { host: 'AUFA', tanggal: '2026-07-08', jamMulai: '19.00', sesi: 'SESI 2', gmv: 715000, target: '0 - 2JT', timestamp: '20/07/2026 11:14:22' },
  { host: 'AUFA', tanggal: '2026-07-09', jamMulai: '05.00', sesi: 'SESI 1', gmv: 366000, target: '0 - 2JT', timestamp: '20/07/2026 11:14:56' },
  { host: 'AUFA', tanggal: '2026-07-09', jamMulai: '11.00', sesi: 'SESI 2', gmv: 797000, target: '0 - 2JT', timestamp: '20/07/2026 11:15:31' },
  { host: 'AUFA', tanggal: '2026-07-10', jamMulai: '05.00', sesi: 'SESI 1', gmv: 502000, target: '0 - 2JT', timestamp: '20/07/2026 11:16:05' },
  { host: 'AUFA', tanggal: '2026-07-10', jamMulai: '09.00', sesi: 'SESI 2', gmv: 1100000, target: '0 - 2JT', timestamp: '20/07/2026 11:16:46' },
  { host: 'AUFA', tanggal: '2026-07-11', jamMulai: '09.00', sesi: 'SESI 1', gmv: 491000, target: '0 - 2JT', timestamp: '20/07/2026 11:17:16' },
  { host: 'AUFA', tanggal: '2026-07-11', jamMulai: '17.00', sesi: 'SESI 2', gmv: 882000, target: '0 - 2JT', timestamp: '20/07/2026 11:17:58' },
  { host: 'AUFA', tanggal: '2026-07-12', jamMulai: '11.00', sesi: 'SESI 1', gmv: 473000, target: '0 - 2JT', timestamp: '20/07/2026 11:18:37' },
  { host: 'AUFA', tanggal: '2026-07-12', jamMulai: '17.00', sesi: 'SESI 2', gmv: 1000000, target: '0 - 2JT', timestamp: '20/07/2026 11:19:03' },
  { host: 'AUFA', tanggal: '2026-07-13', jamMulai: '09.00', sesi: 'SESI 1', gmv: 133000, target: '0 - 2JT', timestamp: '20/07/2026 11:19:31' },
  { host: 'AUFA', tanggal: '2026-07-13', jamMulai: '17.00', sesi: 'SESI 2', gmv: 470000, target: '0 - 2JT', timestamp: '20/07/2026 11:20:02' },
  { host: 'AUFA', tanggal: '2026-07-14', jamMulai: '11.00', sesi: 'SESI 1', gmv: 365000, target: '0 - 2JT', timestamp: '20/07/2026 11:20:40' },
  { host: 'AUFA', tanggal: '2026-07-14', jamMulai: '17.00', sesi: 'SESI 2', gmv: 920000, target: '0 - 2JT', timestamp: '20/07/2026 11:21:08' },
  { host: 'AUFA', tanggal: '2026-07-15', jamMulai: '09.00', sesi: 'SESI 1', gmv: 729000, target: '0 - 2JT', timestamp: '20/07/2026 11:21:36' },
  { host: 'AUFA', tanggal: '2026-07-15', jamMulai: '17.00', sesi: 'SESI 2', gmv: 744000, target: '0 - 2JT', timestamp: '20/07/2026 11:22:08' },
  { host: 'AUFA', tanggal: '2026-07-16', jamMulai: '07.00', sesi: 'SESI 1', gmv: 364000, target: '0 - 2JT', timestamp: '20/07/2026 11:22:41' },
  { host: 'AUFA', tanggal: '2026-07-16', jamMulai: '17.00', sesi: 'SESI 2', gmv: 155000, target: '0 - 2JT', timestamp: '20/07/2026 11:23:11' },
  { host: 'AUFA', tanggal: '2026-07-17', jamMulai: '07.00', sesi: 'SESI 1', gmv: 409000, target: '0 - 2JT', timestamp: '20/07/2026 11:23:47' },
  { host: 'AUFA', tanggal: '2026-07-17', jamMulai: '17.00', sesi: 'SESI 2', gmv: 524000, target: '0 - 2JT', timestamp: '20/07/2026 11:24:15' },
  { host: 'AUFA', tanggal: '2026-07-18', jamMulai: '09.00', sesi: 'SESI 1', gmv: 404000, target: '0 - 2JT', timestamp: '20/07/2026 11:24:55' },
  { host: 'AUFA', tanggal: '2026-07-18', jamMulai: '17.00', sesi: 'SESI 2', gmv: 408000, target: '0 - 2JT', timestamp: '20/07/2026 11:25:25' },
  { host: 'AUFA', tanggal: '2026-07-19', jamMulai: '09.00', sesi: 'SESI 1', gmv: 484000, target: '0 - 2JT', timestamp: '20/07/2026 11:26:01' },
  { host: 'AUFA', tanggal: '2026-07-19', jamMulai: '17.00', sesi: 'SESI 2', gmv: 547000, target: '0 - 2JT', timestamp: '20/07/2026 11:26:43' },
  { host: 'AUFA', tanggal: '2026-07-20', jamMulai: '09.00', sesi: 'SESI 1', gmv: 256000, target: '0 - 2JT', timestamp: '20/07/2026 11:27:09' },
  { host: 'ALIN', tanggal: '2026-07-20', jamMulai: '07.00', sesi: 'SESI 1', gmv: 274000, target: '0 - 2JT', timestamp: '20/07/2026 15:05:22' },
  { host: 'ALIN', tanggal: '2026-07-20', jamMulai: '13.00', sesi: 'SESI 2', gmv: 553000, target: '0 - 2JT', timestamp: '20/07/2026 15:06:02' },
  { host: 'APRIL', tanggal: '2026-07-17', jamMulai: '05.00', sesi: 'SESI 1', gmv: 201000, target: '0 - 2JT', timestamp: '21/07/2026 21:47:27' },
  { host: 'APRIL', tanggal: '2026-07-17', jamMulai: '11.00', sesi: 'SESI 2', gmv: 659000, target: '0 - 2JT', timestamp: '21/07/2026 21:48:41' },
  { host: 'APRIL', tanggal: '2026-07-18', jamMulai: '11.00', sesi: 'SESI 1', gmv: 178000, target: '0 - 2JT', timestamp: '21/07/2026 21:49:30' },
  { host: 'APRIL', tanggal: '2026-07-18', jamMulai: '15.00', sesi: 'SESI 2', gmv: 367000, target: '0 - 2JT', timestamp: '21/07/2026 21:50:02' },
  { host: 'APRIL', tanggal: '2026-07-19', jamMulai: '05.00', sesi: 'SESI 1', gmv: 153000, target: '0 - 2JT', timestamp: '21/07/2026 21:50:49' },
  { host: 'APRIL', tanggal: '2026-07-19', jamMulai: '09.00', sesi: 'SESI 2', gmv: 198000, target: '0 - 2JT', timestamp: '21/07/2026 21:51:26' },
  { host: 'APRIL', tanggal: '2026-07-20', jamMulai: '05.00', sesi: 'SESI 1', gmv: 101000, target: '0 - 2JT', timestamp: '21/07/2026 21:51:56' },
  { host: 'APRIL', tanggal: '2026-07-20', jamMulai: '11.00', sesi: 'SESI 2', gmv: 476000, target: '0 - 2JT', timestamp: '21/07/2026 21:52:28' },
  { host: 'APRIL', tanggal: '2026-07-21', jamMulai: '05.00', sesi: 'SESI 1', gmv: 411000, target: '0 - 2JT', timestamp: '21/07/2026 21:52:51' },
  { host: 'APRIL', tanggal: '2026-07-21', jamMulai: '11.00', sesi: 'SESI 2', gmv: 52000, target: '0 - 2JT', timestamp: '21/07/2026 21:53:15' },
  { host: 'AUFA', tanggal: '2026-07-20', jamMulai: '17.00', sesi: 'SESI 2', gmv: 542000, target: '0 - 2JT', timestamp: '22/07/2026 13:05:07' },
  { host: 'AUFA', tanggal: '2026-07-21', jamMulai: '09.00', sesi: 'SESI 1', gmv: 503000, target: '0 - 2JT', timestamp: '22/07/2026 13:05:39' },
  { host: 'AUFA', tanggal: '2026-07-21', jamMulai: '17.00', sesi: 'SESI 2', gmv: 402000, target: '0 - 2JT', timestamp: '22/07/2026 13:06:08' },
  { host: 'AUFA', tanggal: '2026-07-22', jamMulai: '05.00', sesi: 'SESI 1', gmv: 260000, target: '0 - 2JT', timestamp: '22/07/2026 13:06:36' },
  { host: 'ALIN', tanggal: '2026-07-21', jamMulai: '07.00', sesi: 'SESI 1', gmv: 532000, target: '0 - 2JT', timestamp: '22/07/2026 14:29:47' },
  { host: 'ALIN', tanggal: '2026-07-21', jamMulai: '13.00', sesi: 'SESI 2', gmv: 157000, target: '0 - 2JT', timestamp: '22/07/2026 14:30:12' },
  { host: 'ALIN', tanggal: '2026-07-22', jamMulai: '07.00', sesi: 'SESI 1', gmv: 613000, target: '0 - 2JT', timestamp: '22/07/2026 14:30:33' },
  { host: 'ALIN', tanggal: '2026-07-22', jamMulai: '13.00', sesi: 'SESI 2', gmv: 535000, target: '0 - 2JT', timestamp: '22/07/2026 15:05:57' },
];

function parseTargetRange(t) {
  const nums = (t.match(/\d+/g) || []).map(Number);
  if (nums.length < 2) return [0, 0];
  return [nums[0] * 1_000_000, nums[1] * 1_000_000];
}

function formatRupiah(n) {
  return "Rp " + Number(n || 0).toLocaleString("id-ID");
}

function nowTimestamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}
function formatTanggalDMY(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

function excelSerialToISO(serial) {
  const utcDays = Math.floor(serial - 25569);
  const date = new Date(utcDays * 86400 * 1000);
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, "0");
  const d = String(date.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}
function parseTanggalFleksibel(val) {
  if (typeof val === "number") return excelSerialToISO(val);
  if (typeof val === "string") {
    let m = val.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
    m = val.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  }
  return "";
}
function parseGmvFleksibel(val) {
  if (typeof val === "number") return Math.round(val);
  if (typeof val === "string") {
    const cleaned = val.replace(/[^\d]/g, "");
    return cleaned ? parseInt(cleaned, 10) : 0;
  }
  return 0;
}
function cariKolom(row, keywords) {
  const keys = Object.keys(row);
  for (const k of keys) {
    const norm = k.toLowerCase().replace(/[^a-z0-9]/g, "");
    if (keywords.some((kw) => norm.includes(kw))) return k;
  }
  return null;
}
function parseBarisExcel(rows) {
  const hasil = [];
  for (const row of rows) {
    const hostKey = cariKolom(row, ["host", "nama"]);
    const tglKey = cariKolom(row, ["tanggal", "tgl", "date"]);
    const jamKey = cariKolom(row, ["jam"]);
    const sesiKey = cariKolom(row, ["sesi"]);
    const gmvKey = cariKolom(row, ["gmv", "perolehan"]);
    const targetKey = cariKolom(row, ["target"]);
    const tsKey = cariKolom(row, ["timestamp"]);
    if (!hostKey || !tglKey || !gmvKey) continue;

    const host = String(row[hostKey] || "").trim().toUpperCase();
    const tanggal = parseTanggalFleksibel(row[tglKey]);
    if (!host || !tanggal) continue;

    let sesi = sesiKey ? String(row[sesiKey]).trim().toUpperCase() : "SESI 1";
    if (sesi && !sesi.startsWith("SESI")) sesi = `SESI ${sesi}`;

    hasil.push({
      host,
      tanggal,
      jamMulai: jamKey ? String(row[jamKey]).trim() : "",
      sesi,
      gmv: parseGmvFleksibel(row[gmvKey]),
      target: targetKey && row[targetKey] ? String(row[targetKey]) : "0 - 2JT",
      timestamp: tsKey ? String(row[tsKey]) : nowTimestamp(),
    });
  }
  return hasil;
}

async function loadShared(key, fallback) {
  try {
    const { data, error } = await supabase.from("kv_store").select("value").eq("key", key).maybeSingle();
    if (error || !data) return fallback;
    return data.value;
  } catch {
    return fallback;
  }
}
async function saveShared(key, value) {
  try {
    await supabase.from("kv_store").upsert({ key, value, updated_at: new Date().toISOString() });
  } catch (e) {
    console.error("storage error", e);
  }
}

export default function App() {
  const [tab, setTab] = useState("input");
  const [hosts, setHosts] = useState(DEFAULT_HOSTS);
  const [entries, setEntries] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showHostSettings, setShowHostSettings] = useState(false);
  const [newHostName, setNewHostName] = useState("");
  const [adminPin, setAdminPin] = useState("1234");
  const [isAdmin, setIsAdmin] = useState(false);
  const [showPinPrompt, setShowPinPrompt] = useState(false);
  const [pinInput, setPinInput] = useState("");
  const [pinError, setPinError] = useState(false);
  const [newPin, setNewPin] = useState("");
  const [importMsg, setImportMsg] = useState("");
  const [excelMsg, setExcelMsg] = useState("");
  const fileInputRef = useRef(null);
  const [expandedHosts, setExpandedHosts] = useState({});
  function toggleHostExpand(host) {
    setExpandedHosts((prev) => ({ ...prev, [host]: !prev[host] }));
  }

  const [form, setForm] = useState({
    host: "", tanggal: "", jamMulai: "", sesi: "", gmv: "", target: "",
  });
  const [justSaved, setJustSaved] = useState(false);

  useEffect(() => {
    (async () => {
      const [h, e, pin] = await Promise.all([
        loadShared("live-hosts", DEFAULT_HOSTS),
        loadShared("live-entries", []),
        loadShared("live-admin-pin", "1234"),
      ]);
      setHosts(h);
      setEntries(e);
      setAdminPin(pin);
      setLoading(false);
    })();
  }, []);

  const isDuplicate = useMemo(() => {
    if (!form.host || !form.tanggal || !form.sesi) return false;
    return entries.some((e) => e.host === form.host && e.tanggal === form.tanggal && e.sesi === form.sesi);
  }, [form.host, form.tanggal, form.sesi, entries]);

  const canSubmit = form.host && form.tanggal && form.jamMulai && form.sesi && form.gmv && form.target && !isDuplicate;

  async function submitEntry() {
    if (!canSubmit) return;
    const entry = { id: Date.now().toString(), ...form, gmv: Number(form.gmv), timestamp: nowTimestamp() };
    const next = [entry, ...entries];
    setEntries(next);
    await saveShared("live-entries", next);
    setForm({ host: "", tanggal: "", jamMulai: "", sesi: "", gmv: "", target: "" });
    setJustSaved(true);
    setTimeout(() => setJustSaved(false), 2000);
  }

  async function deleteEntry(id) {
    if (!isAdmin) return;
    const next = entries.filter((e) => e.id !== id);
    setEntries(next);
    await saveShared("live-entries", next);
  }

  async function addHost() {
    const name = newHostName.trim().toUpperCase();
    if (!name || hosts.includes(name)) return;
    const next = [...hosts, name];
    setHosts(next);
    await saveShared("live-hosts", next);
    setNewHostName("");
  }
  async function removeHost(name) {
    const next = hosts.filter((h) => h !== name);
    setHosts(next);
    await saveShared("live-hosts", next);
  }

  function tryUnlock() {
    if (pinInput === adminPin) {
      setIsAdmin(true);
      setShowPinPrompt(false);
      setPinInput("");
      setPinError(false);
    } else {
      setPinError(true);
    }
  }

  async function changePin() {
    const p = newPin.trim();
    if (!p) return;
    setAdminPin(p);
    await saveShared("live-admin-pin", p);
    setNewPin("");
  }

  async function importJuli2026() {
    const existingKeys = new Set(entries.map((e) => `${e.host}|${e.tanggal}|${e.sesi}`));
    const toAdd = IMPORT_DATA_JULI_2026.filter((r) => !existingKeys.has(`${r.host}|${r.tanggal}|${r.sesi}`)).map((r, i) => ({
      id: `import-${r.host}-${r.tanggal}-${r.sesi}-${i}`.replace(/\s/g, ""),
      ...r,
    }));
    if (toAdd.length === 0) {
      setImportMsg("Semua data sudah pernah diimpor, tidak ada yang ditambahkan.");
      return;
    }
    const next = [...entries, ...toAdd];
    setEntries(next);
    await saveShared("live-entries", next);
    setImportMsg(`${toAdd.length} laporan berhasil diimpor (${IMPORT_DATA_JULI_2026.length - toAdd.length} dilewati karena sudah ada).`);
  }

  async function handleExcelFile(fileEvent) {
    const file = fileEvent.target.files?.[0];
    fileEvent.target.value = "";
    if (!file) return;
    setExcelMsg("Membaca file…");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      let allRows = [];
      wb.SheetNames.forEach((name) => {
        const sheet = wb.Sheets[name];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        allRows = allRows.concat(rows);
      });
      const parsed = parseBarisExcel(allRows);
      if (parsed.length === 0) {
        setExcelMsg("Tidak ada baris yang bisa dibaca. Pastikan ada kolom Nama Host, Tanggal, dan GMV.");
        return;
      }
      const newHosts = [...hosts];
      parsed.forEach((p) => { if (!newHosts.includes(p.host)) newHosts.push(p.host); });
      if (newHosts.length !== hosts.length) {
        setHosts(newHosts);
        await saveShared("live-hosts", newHosts);
      }
      const existingKeys = new Set(entries.map((e) => `${e.host}|${e.tanggal}|${e.sesi}`));
      const toAdd = parsed
        .filter((p) => !existingKeys.has(`${p.host}|${p.tanggal}|${p.sesi}`))
        .map((p, i) => ({ id: `xls-${Date.now()}-${i}`, ...p }));
      if (toAdd.length === 0) {
        setExcelMsg("Semua baris di file ini sudah pernah ada, tidak ada yang ditambahkan.");
        return;
      }
      const next = [...entries, ...toAdd];
      setEntries(next);
      await saveShared("live-entries", next);
      setExcelMsg(`${toAdd.length} laporan berhasil diimpor dari Excel (${parsed.length - toAdd.length} dilewati karena duplikat).`);
    } catch (err) {
      console.error(err);
      setExcelMsg("Gagal membaca file. Pastikan formatnya .xlsx atau .xls.");
    }
  }

  const months = useMemo(() => {
    const set = new Set(entries.map((e) => e.tanggal?.slice(0, 7)).filter(Boolean));
    return Array.from(set).sort().reverse();
  }, [entries]);
  const [monthFilter, setMonthFilter] = useState("");
  useEffect(() => { if (months.length && !monthFilter) setMonthFilter(months[0]); }, [months]);

  const filtered = useMemo(
    () => entries.filter((e) => !monthFilter || e.tanggal?.startsWith(monthFilter)),
    [entries, monthFilter]
  );

  const perHost = useMemo(() => {
    const map = {};
    hosts.forEach((h) => (map[h] = { host: h, total: 0, sesi: 0, tercapai: 0, gajiPokok: 0, bonus: 0, totalGaji: 0, hariSet: new Set() }));
    filtered.forEach((e) => {
      if (!map[e.host]) map[e.host] = { host: e.host, total: 0, sesi: 0, tercapai: 0, gajiPokok: 0, bonus: 0, totalGaji: 0, hariSet: new Set() };
      map[e.host].total += e.gmv;
      map[e.host].sesi += 1;
      map[e.host].gajiPokok += BASE_UPAH;
      map[e.host].bonus += hitungBonus(e.gmv);
      map[e.host].totalGaji += hitungTotalGaji(e.gmv);
      if (e.gmv >= 2_000_000) map[e.host].tercapai += 1;
      if (e.tanggal) map[e.host].hariSet.add(e.tanggal);
    });
    return Object.values(map)
      .map((p) => ({ ...p, hariKerja: p.hariSet.size }))
      .sort((a, b) => b.total - a.total);
  }, [filtered, hosts]);

  const topPerformer = perHost.length > 0 && perHost[0].total > 0 ? perHost[0].host : null;

  const dailyChartData = useMemo(() => {
    if (!monthFilter) return [];
    const [y, m] = monthFilter.split("-").map(Number);
    if (!y || !m) return [];
    const daysInMonth = new Date(y, m, 0).getDate();
    const arr = [];
    for (let d = 1; d <= daysInMonth; d++) {
      const row = { day: d };
      hosts.forEach((h) => { row[h] = 0; });
      arr.push(row);
    }
    filtered.forEach((e) => {
      const day = parseInt((e.tanggal || "").slice(8, 10), 10);
      if (day >= 1 && day <= daysInMonth) {
        arr[day - 1][e.host] = (arr[day - 1][e.host] || 0) + e.gmv;
      }
    });
    return arr;
  }, [filtered, monthFilter, hosts]);

  const detailByHost = useMemo(() => {
    const map = {};
    hosts.forEach((h) => (map[h] = []));
    filtered.forEach((e) => {
      if (!map[e.host]) map[e.host] = [];
      map[e.host].push(e);
    });
    Object.keys(map).forEach((h) => {
      map[h].sort((a, b) => (b.tanggal + b.jamMulai).localeCompare(a.tanggal + a.jamMulai));
    });
    return map;
  }, [filtered, hosts]);

  const grandTotal = filtered.reduce((s, e) => s + e.gmv, 0);
  const grandGajiPokok = filtered.length * BASE_UPAH;
  const grandBonus = filtered.reduce((s, e) => s + hitungBonus(e.gmv), 0);
  const grandTotalGaji = grandGajiPokok + grandBonus;

  function exportExcel() {
    const detailRows = filtered.map((e) => {
      return {
        Timestamp: e.timestamp || "",
        "NAMA HOST": e.host,
        TANGGAL: formatTanggalDMY(e.tanggal),
        "JAM MULAI": e.jamMulai,
        SESI: e.sesi,
        "LAPORAN PEROLEHAN GMV TIKTOK": e.gmv,
        TARGET: e.target,
        STATUS: e.gmv >= 2_000_000 ? "Tercapai" : "Belum",
        GAJI: BASE_UPAH,
        BONUS: hitungBonus(e.gmv),
        "TOTAL GAJI": hitungTotalGaji(e.gmv),
      };
    });
    const summaryRows = perHost.map((p) => ({
      Host: p.host,
      "Total GMV": p.total,
      "Jumlah Sesi": p.sesi,
      "Sesi Tercapai": p.tercapai,
      Gaji: p.gajiPokok,
      Bonus: p.bonus,
      "Total Gaji": p.totalGaji,
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "Rekap Host");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detailRows), "Detail Laporan");
    XLSX.writeFile(wb, `laporan-live-${monthFilter || "semua"}.xlsx`);
  }

  return (
    <div style={{ minHeight: "100vh", background: "#0a0a0f", backgroundImage: "radial-gradient(circle at 15% 0%, rgba(37,244,238,0.15), transparent 45%), radial-gradient(circle at 90% 15%, rgba(254,44,85,0.18), transparent 40%)", fontFamily: "Inter, sans-serif", color: "#F5F1E8", paddingBottom: 40 }}>
      <style>{FONT_STYLE}</style>

      <header style={{ padding: "28px 20px 20px", borderBottom: "1px solid #24242f" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <Radio size={20} color={ACCENT} />
            <span style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12, letterSpacing: 2, color: CYAN, textTransform: "uppercase" }}>
              Live GMV Tracker
            </span>
          </div>
          <button
            onClick={() => (isAdmin ? setIsAdmin(false) : setShowPinPrompt(true))}
            title={isAdmin ? "Kunci mode admin" : "Buka mode admin"}
            style={{ background: "none", border: "none", color: isAdmin ? GOOD : "#8a8a9a", cursor: "pointer" }}>
            {isAdmin ? <Unlock size={18} /> : <Lock size={18} />}
          </button>
        </div>
        <h1 style={{
          fontFamily: "Fredoka, sans-serif", fontWeight: 600, fontSize: 30, margin: "10px 0 0",
          background: `linear-gradient(90deg, ${CYAN}, ${ACCENT})`,
          WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text",
        }}>
          Perolehan Live Streaming Faradisaaaa
        </h1>
        <div style={{ display: "flex", gap: 8, marginTop: 18 }}>
          {["input", "rekap"].map((t) => (
            <button key={t} onClick={() => setTab(t)}
              style={{
                flex: 1, padding: "10px 0", borderRadius: 10, border: "none", cursor: "pointer",
                fontFamily: "Inter", fontWeight: 600, fontSize: 14,
                background: tab === t ? `linear-gradient(135deg, ${ACCENT}, ${CYAN})` : "#1c1c26",
                color: tab === t ? "#14141c" : "#F5F1E8",
              }}>
              {t === "input" ? "Input Laporan" : "Rekap"}
            </button>
          ))}
        </div>
      </header>

      {loading ? (
        <div style={{ padding: 40, textAlign: "center", opacity: 0.6 }}>Memuat data…</div>
      ) : tab === "input" ? (
        <div style={{ padding: 20, maxWidth: 480, margin: "0 auto" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
            <label style={labelStyle}>NAMA HOST</label>
            {isAdmin && (
              <button onClick={() => setShowHostSettings(true)} style={{ background: "none", border: "none", color: "#8a8a9a", cursor: "pointer" }}>
                <Settings2 size={16} />
              </button>
            )}
          </div>
          <select value={form.host} onChange={(e) => setForm({ ...form, host: e.target.value })} style={selectStyle}>
            <option value="">Pilih host</option>
            {hosts.map((h) => <option key={h} value={h}>{h}</option>)}
          </select>

          <label style={labelStyle}>TANGGAL</label>
          <input type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} style={selectStyle} />

          <label style={labelStyle}>JAM MULAI</label>
          <select value={form.jamMulai} onChange={(e) => setForm({ ...form, jamMulai: e.target.value })} style={selectStyle}>
            <option value="">Pilih jam</option>
            {JAM_OPTIONS.map((j) => <option key={j} value={j}>{j}</option>)}
          </select>

          <label style={labelStyle}>SESI</label>
          <select value={form.sesi} onChange={(e) => setForm({ ...form, sesi: e.target.value })} style={selectStyle}>
            <option value="">Pilih sesi</option>
            {SESI_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>

          <label style={labelStyle}>LAPORAN PEROLEHAN GMV TIKTOK</label>
          <input type="number" placeholder="misal 8500000" value={form.gmv}
            onChange={(e) => setForm({ ...form, gmv: e.target.value })} style={{ ...selectStyle, fontFamily: "'JetBrains Mono', monospace" }} />

          <label style={labelStyle}>TARGET</label>
          <select value={form.target} onChange={(e) => setForm({ ...form, target: e.target.value })} style={selectStyle}>
            <option value="">Pilih target</option>
            {TARGET_OPTIONS.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>

          <button onClick={submitEntry} disabled={!canSubmit}
            style={{
              width: "100%", marginTop: 20, padding: "14px 0", borderRadius: 10, border: "none",
              background: canSubmit ? `linear-gradient(135deg, ${ACCENT}, ${CYAN})` : "#2a2a35", color: canSubmit ? "#0a0a0f" : "#77778a",
              fontWeight: 700, fontSize: 15, cursor: canSubmit ? "pointer" : "not-allowed",
              display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
            }}>
            <Plus size={18} /> Simpan Laporan
          </button>
          {isDuplicate && (
            <div style={{ marginTop: 10, textAlign: "center", color: BAD, fontSize: 13, fontWeight: 600 }}>
              ⚠️ Laporan {form.host} tanggal {formatTanggalDMY(form.tanggal)} {form.sesi} sudah pernah diinput.
            </div>
          )}
          {justSaved && <div style={{ marginTop: 10, textAlign: "center", color: GOOD, fontSize: 13, fontWeight: 600 }}>✨ Laporan tersimpan!</div>}
        </div>
      ) : (
        <div style={{ padding: 20, maxWidth: 640, margin: "0 auto" }}>
          <div style={{ display: "flex", gap: 8 }}>
            <select value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)} style={{ ...selectStyle, maxWidth: 220 }}>
              {months.length === 0 && <option value="">Belum ada data</option>}
              {months.map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
            <button onClick={exportExcel} disabled={filtered.length === 0}
              style={{
                display: "flex", alignItems: "center", gap: 6, whiteSpace: "nowrap",
                padding: "0 16px", borderRadius: 10, border: "1px solid #2a2a36",
                background: filtered.length === 0 ? "#1c1c26" : GOOD,
                color: filtered.length === 0 ? "#77778a" : "#14141c",
                fontWeight: 600, fontSize: 13, cursor: filtered.length === 0 ? "not-allowed" : "pointer",
              }}>
              <FileDown size={16} /> Excel
            </button>
          </div>

          <div style={{ display: "flex", gap: 10, margin: "16px 0" }}>
            <div style={cardStyle}>
              <div style={{ fontSize: 12, color: "#8a8a9a" }}>TOTAL GMV</div>
              <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 18, fontWeight: 700, color: ACCENT }}>{formatRupiah(grandTotal)}</div>
            </div>
            <div style={cardStyle}>
              <div style={{ fontSize: 12, color: "#8a8a9a" }}>TOTAL SESI</div>
              <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 18, fontWeight: 700 }}>{filtered.length}</div>
            </div>
          </div>
          {isAdmin && (
            <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
              <div style={cardStyle}>
                <div style={{ fontSize: 12, color: "#8a8a9a" }}>GAJI</div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 16, fontWeight: 700 }}>{formatRupiah(grandGajiPokok)}</div>
              </div>
              <div style={cardStyle}>
                <div style={{ fontSize: 12, color: "#8a8a9a" }}>BONUS</div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 16, fontWeight: 700, color: ACCENT }}>{formatRupiah(grandBonus)}</div>
              </div>
              <div style={cardStyle}>
                <div style={{ fontSize: 12, color: "#8a8a9a" }}>TOTAL GAJI</div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 16, fontWeight: 700, color: GOOD }}>{formatRupiah(grandTotalGaji)}</div>
              </div>
            </div>
          )}

          <div style={{ height: 200, marginBottom: 20 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={perHost}>
                <CartesianGrid strokeDasharray="3 3" stroke="#2a2a36" />
                <XAxis dataKey="host" tick={{ fill: "#8a8a9a", fontSize: 11 }} />
                <YAxis tick={{ fill: "#8a8a9a", fontSize: 10 }} tickFormatter={(v) => `${v / 1e6}jt`} />
                <Tooltip formatter={(v) => formatRupiah(v)} contentStyle={{ background: "#20202b", border: "none", borderRadius: 8 }} />
                <Bar dataKey="total" fill={CYAN} radius={[6, 6, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>

          <h3 style={{ fontFamily: "Fredoka, sans-serif", fontSize: 16, marginBottom: 10, color: "#8a8a9a" }}>Tren GMV Harian</h3>
          <div style={{ height: 220, marginBottom: 20 }}>
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={dailyChartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#2a2a36" />
                <XAxis dataKey="day" tick={{ fill: "#8a8a9a", fontSize: 10 }} />
                <YAxis tick={{ fill: "#8a8a9a", fontSize: 10 }} tickFormatter={(v) => `${v / 1e6}jt`} />
                <Tooltip formatter={(v) => formatRupiah(v)} labelFormatter={(d) => `Tanggal ${d}`} contentStyle={{ background: "#20202b", border: "none", borderRadius: 8 }} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                {hosts.map((h, i) => (
                  <Line key={h} type="monotone" dataKey={h} stroke={HOST_COLORS[i % HOST_COLORS.length]} strokeWidth={2} dot={false} />
                ))}
              </LineChart>
            </ResponsiveContainer>
          </div>

          <h3 style={{ fontFamily: "Fredoka, sans-serif", fontSize: 18, marginBottom: 10 }}>Rekap per Host</h3>
          {perHost.map((p) => (
            <div key={p.host} style={{ ...cardStyle, display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div>
                <div style={{ fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
                  {p.host}
                  {topPerformer === p.host && <Crown size={14} color={BAD} />}
                </div>
                <div style={{ fontSize: 12, color: "#8a8a9a" }}>{p.hariKerja} hari kerja · {p.sesi} sesi · {p.tercapai} tercapai target</div>
                {isAdmin && (
                  <div style={{ fontSize: 11, color: "#8a8a9a", marginTop: 2 }}>
                    Gaji {formatRupiah(p.gajiPokok)} + Bonus {formatRupiah(p.bonus)}
                  </div>
                )}
              </div>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontWeight: 700, color: ACCENT }}>{formatRupiah(p.total)}</div>
                {isAdmin && (
                  <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12, color: GOOD }}>Total: {formatRupiah(p.totalGaji)}</div>
                )}
              </div>
            </div>
          ))}

          <h3 style={{ fontFamily: "Fredoka, sans-serif", fontSize: 18, margin: "20px 0 10px" }}>Detail Laporan</h3>
          {filtered.length === 0 && <div style={{ opacity: 0.5, fontSize: 14 }}>Belum ada laporan bulan ini.</div>}
          {hosts.filter((h) => (detailByHost[h] || []).length > 0).map((host) => {
            const hostEntries = detailByHost[host] || [];
            const isOpen = !!expandedHosts[host];
            return (
              <div key={host} style={{ ...cardStyle, marginBottom: 8, padding: 0, overflow: "hidden" }}>
                <button onClick={() => toggleHostExpand(host)}
                  style={{ width: "100%", background: "none", border: "none", color: "#F5F1E8", cursor: "pointer", padding: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
                    {isOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                    {host}
                  </span>
                  <span style={{ fontSize: 12, color: "#8a8a9a" }}>{hostEntries.length} laporan</span>
                </button>
                {isOpen && (
                  <div style={{ padding: "0 14px 14px" }}>
                    {hostEntries.map((e) => {
                      const tercapai = e.gmv >= 2_000_000;
                      return (
                        <div key={e.id} style={{ borderTop: "1px solid #24242f", paddingTop: 10, marginTop: 10 }}>
                          <div style={{ display: "flex", justifyContent: "space-between" }}>
                            <div style={{ fontSize: 13, color: "#8a8a9a" }}>{formatTanggalDMY(e.tanggal)} · {e.jamMulai} · {e.sesi}</div>
                            {isAdmin && (
                              <button onClick={() => deleteEntry(e.id)} style={{ background: "none", border: "none", color: "#8a8a9a", cursor: "pointer" }}>
                                <Trash2 size={14} />
                              </button>
                            )}
                          </div>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 6, flexWrap: "wrap" }}>
                            <span style={{ fontFamily: "'JetBrains Mono', monospace", fontWeight: 700 }}>{formatRupiah(e.gmv)}</span>
                            <span style={{ fontSize: 11, padding: "2px 8px", borderRadius: 20, background: tercapai ? "rgba(37,244,238,0.15)" : "rgba(255,184,0,0.15)", color: tercapai ? GOOD : BAD, fontWeight: 600 }}>
                              {tercapai ? "🎉 Tercapai" : "Belum"}
                            </span>
                          </div>
                          {isAdmin && (
                            <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12, color: "#8a8a9a", marginTop: 4 }}>
                              Gaji {formatRupiah(BASE_UPAH)} + Bonus {formatRupiah(hitungBonus(e.gmv))} = <span style={{ color: GOOD }}>{formatRupiah(hitungTotalGaji(e.gmv))}</span>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {showHostSettings && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", zIndex: 10 }}>
          <div style={{ background: "#1c1c26", width: "100%", borderRadius: "16px 16px 0 0", padding: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <h3 style={{ fontFamily: "Fredoka, sans-serif", fontSize: 18, margin: 0 }}>Kelola Host</h3>
              <button onClick={() => setShowHostSettings(false)} style={{ background: "none", border: "none", color: "#F5F1E8", cursor: "pointer" }}>
                <X size={20} />
              </button>
            </div>
            {hosts.map((h) => (
              <div key={h} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid #2a2a36" }}>
                <span>{h}</span>
                <button onClick={() => removeHost(h)} style={{ background: "none", border: "none", color: BAD, cursor: "pointer" }}>
                  <Trash2 size={16} />
                </button>
              </div>
            ))}
            <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
              <input value={newHostName} onChange={(e) => setNewHostName(e.target.value)} placeholder="Nama host baru"
                style={{ ...selectStyle, margin: 0, flex: 1 }} />
              <button onClick={addHost} style={{ background: ACCENT, border: "none", borderRadius: 10, padding: "0 16px", color: "#14141c", fontWeight: 700, cursor: "pointer" }}>
                Tambah
              </button>
            </div>

            <div style={{ marginTop: 20, paddingTop: 16, borderTop: "1px solid #2a2a36" }}>
              <label style={labelStyle}>GANTI PIN ADMIN</label>
              <div style={{ display: "flex", gap: 8 }}>
                <input value={newPin} onChange={(e) => setNewPin(e.target.value)} placeholder="PIN baru" type="text"
                  style={{ ...selectStyle, margin: 0, flex: 1 }} />
                <button onClick={changePin} style={{ background: GOOD, border: "none", borderRadius: 10, padding: "0 16px", color: "#14141c", fontWeight: 700, cursor: "pointer" }}>
                  Simpan
                </button>
              </div>
            </div>

            <div style={{ marginTop: 20, paddingTop: 16, borderTop: "1px solid #24242f" }}>
              <label style={labelStyle}>IMPORT DATA LAMA</label>
              <button onClick={importJuli2026}
                style={{ width: "100%", padding: "12px 0", borderRadius: 10, border: "none", background: `linear-gradient(135deg, ${ACCENT}, ${CYAN})`, color: "#0a0a0f", fontWeight: 700, cursor: "pointer" }}>
                Import Data Juli 2026 ({IMPORT_DATA_JULI_2026.length} baris)
              </button>
              {importMsg && <div style={{ marginTop: 8, fontSize: 12, color: "#8a8a9a" }}>{importMsg}</div>}
            </div>

            <div style={{ marginTop: 20, paddingTop: 16, borderTop: "1px solid #24242f" }}>
              <label style={labelStyle}>IMPORT DARI FILE EXCEL KAMU</label>
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleExcelFile} style={{ display: "none" }} />
              <button onClick={() => fileInputRef.current?.click()}
                style={{ width: "100%", padding: "12px 0", borderRadius: 10, border: "1px solid #24242f", background: "#16161f", color: "#F5F1E8", fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
                <FileDown size={16} style={{ transform: "rotate(180deg)" }} /> Pilih File Excel (.xlsx)
              </button>
              <div style={{ fontSize: 11, color: "#8a8a9a", marginTop: 6 }}>
                Kolom yang dikenali otomatis: Nama Host, Tanggal, Jam Mulai, Sesi, GMV, Target — nama kolom fleksibel, tidak harus persis sama.
              </div>
              {excelMsg && <div style={{ marginTop: 8, fontSize: 12, color: "#8a8a9a" }}>{excelMsg}</div>}
            </div>
          </div>
        </div>
      )}

      {showPinPrompt && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 20 }}>
          <div style={{ background: "#1c1c26", borderRadius: 16, padding: 24, width: "85%", maxWidth: 320 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <h3 style={{ fontFamily: "Fredoka, sans-serif", fontSize: 18, margin: 0 }}>Mode Admin</h3>
              <button onClick={() => { setShowPinPrompt(false); setPinInput(""); setPinError(false); }} style={{ background: "none", border: "none", color: "#F5F1E8", cursor: "pointer" }}>
                <X size={20} />
              </button>
            </div>
            <input
              type="password"
              value={pinInput}
              onChange={(e) => { setPinInput(e.target.value); setPinError(false); }}
              placeholder="Masukkan PIN"
              style={selectStyle}
              onKeyDown={(e) => e.key === "Enter" && tryUnlock()}
            />
            {pinError && <div style={{ color: BAD, fontSize: 12, marginTop: 6 }}>PIN salah, coba lagi.</div>}
            <button onClick={tryUnlock}
              style={{ width: "100%", marginTop: 14, padding: "12px 0", borderRadius: 10, border: "none", background: ACCENT, color: "#14141c", fontWeight: 700, cursor: "pointer" }}>
              Buka
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

const labelStyle = { display: "block", fontSize: 11, letterSpacing: 1, color: "#8a8a9a", margin: "16px 0 6px", fontFamily: "'JetBrains Mono', monospace" };
const selectStyle = { width: "100%", padding: "12px 14px", borderRadius: 12, border: "1px solid #24242f", background: "#16161f", color: "#F5F1E8", fontSize: 14, fontFamily: "Inter", outline: "none", boxSizing: "border-box" };
const cardStyle = { flex: 1, background: "#16161f", border: "1px solid #24242f", borderRadius: 14, padding: 14 };
