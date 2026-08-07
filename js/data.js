/**
 * 團體作品與個人作品資料。
 *
 * 團體作品：內容（組長簡報的夜市機台介紹＋每人分工頁）由
 * scripts/build_group_pages.py 自動產生成 groups/<id>.html 靜態頁面，
 * 該腳本執行時也會印出每組的「招牌照」「機台照」路徑，首頁卡片縮圖
 * （signSrc／deviceSrc）就是用那兩張照片並排顯示。簡報內容更新後重新
 * 執行腳本、把印出的新路徑貼回這裡即可；缺其中一張時另一張會補滿版面。
 *
 * 個人作品：每位學生的內容（文字＋圖片）是由 scripts/build_student_pages.py
 * 從 assets/8-1|8-2|8-3 底下的簡報自動產生成 students/<key>.html 靜態頁面，
 * 卡片縮圖用的照片存在 assets/media/students/<key>-photo.<ext>（缺照片的學生見
 * STUDENT_PHOTO_EXT 之外的 key，會 fallback 成預設頭像）。
 * 簡報內容更新後重新執行這兩支腳本即可，不需要手動改這裡的 pageSrc 邏輯。
 */

const GROUPS = [
  { id: 'a', name: 'A組', members: ['蔡翔宇', '詹柏軒', '蔡丞皓'], leader: '蔡丞皓', machineName: '射企球', signSrc: 'assets/media/groups/a/sign.jpg', deviceSrc: 'assets/media/groups/a/device.png', ready: true },
  { id: 'b', name: 'B組', members: ['黃祈翔', '吳婕語', '石展成', '李心恬'], leader: '黃祈翔', machineName: '67彈珠台', signSrc: 'assets/media/groups/b/sign.png', deviceSrc: 'assets/media/groups/b/device.png', ready: true },
  { id: 'c', name: 'C組', members: ['陳予欣', '陳亮希', '楊初淨', '鍾勻瑨'], leader: '楊初淨', machineName: '我來破壞接棒機', signSrc: 'assets/media/groups/c/sign.png', deviceSrc: 'assets/media/groups/c/device.png', ready: true },
  { id: 'd', name: 'D組', members: ['周以潔', '林為樂', '張宥淇', '周昀希'], leader: '周以潔', machineName: '神射手的挑戰：移動的標靶裝置！！', signSrc: 'assets/media/groups/d/sign.png', deviceSrc: 'assets/media/groups/d/device.png', ready: true },
  { id: 'e', name: 'E組', members: ['曾詠捷', '黃婕霓', '傅蕾棋', '林琍晴'], leader: '傅蕾棋', machineName: '乒乓球機台', signSrc: 'assets/media/groups/e/sign.png', deviceSrc: 'assets/media/groups/e/device.png', ready: true },
  { id: 'f', name: 'F組', members: ['周榆庭', '魏右定', '王儀儼', '周于喆'], leader: '周榆庭', machineName: '彈珠小王子', signSrc: 'assets/media/groups/f/sign.png', deviceSrc: 'assets/media/groups/f/device.png', ready: true },
  { id: 'g', name: 'G組', members: ['許詠晴', '陳維潔', '呂睿恩', '王念安'], leader: '許詠晴', machineName: '彈珠平衡器', signSrc: 'assets/media/groups/g/sign.jpg', deviceSrc: 'assets/media/groups/g/device.png', ready: true },
  { id: 'h', name: 'H組', members: ['蔡寧', '林宇馨', '黃行書'], leader: '蔡寧', machineName: '攀岩地鼠', signSrc: 'assets/media/groups/h/sign.png', deviceSrc: 'assets/media/groups/h/device.png', ready: true },
  { id: 'i', name: 'I組', members: ['張諄楷', '陳定鴻', '劉億承', '李侑磬'], leader: '陳定鴻', machineName: '我的攤位', signSrc: 'assets/media/groups/i/sign.png', deviceSrc: 'assets/media/groups/i/device.png', ready: true },
];

// 每位學生照片的副檔名（key -> 副檔名），目前 34 位都已補齊
const STUDENT_PHOTO_EXT = {
  'a-1': 'png', 'a-2': 'jpg', 'a-3': 'png',
  'b-1': 'png', 'b-2': 'jpg', 'b-3': 'png', 'b-4': 'jpg',
  'c-1': 'png', 'c-2': 'png', 'c-3': 'png', 'c-4': 'jpg',
  'd-1': 'jpg', 'd-2': 'jpg', 'd-3': 'jpg', 'd-4': 'jpg',
  'e-1': 'png', 'e-2': 'jpg', 'e-3': 'png', 'e-4': 'png',
  'f-1': 'jpg', 'f-2': 'png', 'f-3': 'jpg', 'f-4': 'png',
  'g-1': 'jpg', 'g-2': 'jpg', 'g-3': 'jpg', 'g-4': 'jpg',
  'h-1': 'jpg', 'h-2': 'png', 'h-3': 'jpg',
  'i-1': 'jpg', 'i-2': 'jpg', 'i-3': 'jpg', 'i-4': 'jpg',
};

const STUDENTS = GROUPS.flatMap((group) =>
  group.members.map((name, index) => {
    const key = `${group.id}-${index + 1}`;
    const photoExt = STUDENT_PHOTO_EXT[key];
    return {
      name,
      groupId: group.id,
      groupName: group.name,
      key,
      pageSrc: `students/${key}.html`,
      photoSrc: photoExt ? `assets/media/students/${key}-photo.${photoExt}` : '',
      ready: true,
    };
  })
);
