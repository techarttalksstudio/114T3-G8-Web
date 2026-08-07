(function () {
  const groupsGrid = document.getElementById('groups-grid');
  const studentsWrap = document.getElementById('students-wrap');

  function renderGroups() {
    GROUPS.forEach((group) => {
      const card = document.createElement('a');
      card.className = 'card group-card' + (group.ready ? '' : ' is-pending');
      card.href = `groups/${group.id}.html`;

      const media = document.createElement('div');
      media.className = 'group-card__media';
      const photos = group.ready
        ? [
            { src: group.signSrc, alt: group.name + '招牌照片' },
            { src: group.deviceSrc, alt: group.name + '機台照片' },
          ].filter((p) => p.src)
        : [{ src: 'assets/placeholder/photo-placeholder.svg', alt: group.name + '作品縮圖' }];
      photos.forEach((photo) => {
        const img = document.createElement('img');
        img.src = photo.src;
        img.alt = photo.alt;
        if (photos.length === 1) img.className = 'group-card__media-full';
        media.appendChild(img);
      });
      card.appendChild(media);

      const body = document.createElement('div');
      body.className = 'card__body';
      const title = document.createElement('h3');
      title.textContent = group.machineName ? `「${group.machineName}」` : group.name;
      body.appendChild(title);
      if (group.machineName) {
        const groupName = document.createElement('p');
        groupName.className = 'group-card__groupname';
        groupName.textContent = group.name;
        body.appendChild(groupName);
      }
      const members = document.createElement('p');
      members.className = 'group-card__members';
      members.textContent = group.members.join('、');
      body.appendChild(members);
      card.appendChild(body);

      groupsGrid.appendChild(card);
    });
  }

  function renderStudents() {
    GROUPS.forEach((group) => {
      const section = document.createElement('div');
      section.className = 'student-group';

      const heading = document.createElement('h3');
      heading.className = 'student-group__title';
      heading.textContent = group.name;
      section.appendChild(heading);

      const grid = document.createElement('div');
      grid.className = 'student-grid';

      STUDENTS.filter((s) => s.groupId === group.id).forEach((student) => {
        const card = document.createElement('a');
        card.className = 'card student-card';
        card.href = student.pageSrc;

        const media = document.createElement('div');
        media.className = 'student-card__media';
        const photo = document.createElement('img');
        photo.src = student.photoSrc || 'assets/placeholder/photo-placeholder.svg';
        photo.alt = student.name + '的照片';
        media.appendChild(photo);
        card.appendChild(media);

        const name = document.createElement('p');
        name.className = 'student-card__name';
        name.textContent = student.name;
        card.appendChild(name);

        grid.appendChild(card);
      });

      section.appendChild(grid);
      studentsWrap.appendChild(section);
    });
  }

  function initTeacherModal() {
    const modal = document.getElementById('teacher-modal');
    if (!modal) return;

    const open = () => modal.classList.add('is-open');
    const close = () => modal.classList.remove('is-open');

    document.querySelectorAll('[data-teacher-trigger]').forEach((el) => {
      el.addEventListener('click', open);
    });
    document.querySelectorAll('[data-teacher-close]').forEach((el) => {
      el.addEventListener('click', close);
    });
    modal.addEventListener('click', (e) => {
      if (e.target === modal) close();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') close();
    });
  }

  renderGroups();
  renderStudents();
  initTeacherModal();
})();
