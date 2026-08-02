function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

// ==================== ЕДИНЫЙ РОУТЕР ====================
// ВАЖНО: раньше действия на запись (create_booking, register,
// admin_create_booking, admin_cancel_booking, admin_block_computer)
// обрабатывались только в doPost, а фронтенд отправлял их обычным
// fetch() без method:'POST', то есть реальным GET-запросом.
// Из-за этого doGet не находил совпадений и всегда возвращал список
// клубов вместо результата операции — бронирование и админка не
// работали. Теперь и doGet, и doPost используют один и тот же роутер,
// поэтому действие сработает независимо от метода запроса.
function handleRequest(e) {
  const p = (e && e.parameter) ? e.parameter : {};

  try {
    if (p.club_id && !p.action) {
      return getClubDetail(parseInt(p.club_id));
    }
    if (p.action === 'login' && p.phone) {
      return loginPlayer(p.phone.toString());
    }
    if (p.action === 'register' && p.name && p.phone) {
      return registerPlayer(p.name.toString(), p.phone.toString());
    }
    if (p.action === 'my_bookings' && p.player_phone) {
      return getPlayerBookings(p.player_phone.toString());
    }
    if (p.action === 'create_booking') {
      return createBooking(p);
    }
    if (p.action === 'admin_bookings' && p.club_id) {
      return getAdminBookings(parseInt(p.club_id));
    }
    if (p.action === 'admin_login' && p.phone && p.club_id) {
      return adminLogin(p.phone.toString(), parseInt(p.club_id));
    }
    if (p.action === 'admin_create_booking') {
      return adminCreateBooking(p);
    }
    if (p.action === 'admin_cancel_booking') {
      return adminCancelBooking(p);
    }
    if (p.action === 'admin_block_computer') {
      return adminBlockComputer(p);
    }
    if (p.action === 'redeem_bonus') {
      return redeemBonus(p);
    }

    return getClubsList();
  } catch (err) {
    return jsonOut({ error: err.message || 'Server error' });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== БОНУСНАЯ СИСТЕМА ====================
// Настройки начисления/списания бонусов.
// 1 бонус = 1 сум (можно менять BONUS_RATE ниже).
const BONUS_ACCRUAL_PERCENT = 5;   // % от суммы брони начисляется бонусами
const BONUS_MAX_REDEEM_PERCENT = 50; // максимум % от суммы, который можно оплатить бонусами

function addBonusToPlayer(phone, amount) {
  if (!amount || amount <= 0) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const playersSheet = sheet.getSheetByName('players');
  const playersData = playersSheet.getDataRange().getValues();
  const rowIndex = playersData.findIndex(row => row[2] && row[2].toString() === phone);
  if (rowIndex === -1) return;
  const currentBalance = Number(playersData[rowIndex][3]) || 0;
  playersSheet.getRange(rowIndex + 1, 4).setValue(currentBalance + amount);
}

function deductBonusFromPlayer(phone, amount) {
  if (!amount || amount <= 0) return true;
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const playersSheet = sheet.getSheetByName('players');
  const playersData = playersSheet.getDataRange().getValues();
  const rowIndex = playersData.findIndex(row => row[2] && row[2].toString() === phone);
  if (rowIndex === -1) return false;
  const currentBalance = Number(playersData[rowIndex][3]) || 0;
  if (currentBalance < amount) return false;
  playersSheet.getRange(rowIndex + 1, 4).setValue(currentBalance - amount);
  return true;
}

function getPlayerBonusBalance(phone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const playersData = sheet.getSheetByName('players').getDataRange().getValues();
  const player = playersData.slice(1).find(row => row[2] && row[2].toString() === phone);
  return player ? (Number(player[3]) || 0) : 0;
}

// Позволяет фронту заранее списать бонусы при оплате брони (используется
// из processPayment на клиенте перед/вместе с созданием брони).
function redeemBonus(params) {
  const phone = params.player_phone ? params.player_phone.toString() : '';
  const amount = Number(params.amount || 0);
  if (!phone || amount <= 0) {
    return jsonOut({ error: 'Некорректные параметры списания' });
  }
  const ok = deductBonusFromPlayer(phone, amount);
  if (!ok) {
    return jsonOut({ error: 'Недостаточно бонусов' });
  }
  return jsonOut({ success: true, new_balance: getPlayerBonusBalance(phone) });
}

// ==================== КЛУБЫ ====================
function getClubsList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const clubsData = sheet.getSheetByName('clubs').getDataRange().getValues();
  const hallsData = sheet.getSheetByName('halls').getDataRange().getValues();
  const computersData = sheet.getSheetByName('computers').getDataRange().getValues();
  const bookingsData = sheet.getSheetByName('bookings').getDataRange().getValues();

  const clubs = [];
  const now = new Date();

  for (let i = 1; i < clubsData.length; i++) {
    const row = clubsData[i];
    const clubId = row[0];
    if (row[5] !== true) continue;

    const clubHalls = hallsData.filter(h => h[1] === clubId && h[7] === true);
    if (clubHalls.length === 0) continue;

    const minPrice = Math.min(...clubHalls.map(h => h[3]));

    let totalComputers = 0;
    let occupiedNow = 0;

    clubHalls.forEach(hall => {
      const hallId = hall[0];
      const comps = computersData.filter(c => c[1] === hallId && c[3] === 'active');
      totalComputers += comps.length;

      comps.forEach(comp => {
        const activeBooking = bookingsData.find(b =>
          b[4] === comp[0] &&
          b[7] === 'confirmed' &&
          new Date(b[5]) <= now &&
          new Date(new Date(b[5]).getTime() + b[6] * 3600000) > now
        );
        if (activeBooking) occupiedNow++;
      });
    });

    clubs.push({
      id: clubId,
      name: row[1],
      address: row[2],
      phone: row[3],
      photo_url: row[4],
      rating: 4.5 + Math.random() * 0.5,
      freeNow: totalComputers - occupiedNow,
      total: totalComputers,
      minPrice: minPrice
    });
  }

  return jsonOut(clubs);
}

function getClubDetail(clubId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const clubsData = sheet.getSheetByName('clubs').getDataRange().getValues();
  const hallsData = sheet.getSheetByName('halls').getDataRange().getValues();
  const computersData = sheet.getSheetByName('computers').getDataRange().getValues();
  const bookingsData = sheet.getSheetByName('bookings').getDataRange().getValues();

  const clubRow = clubsData.slice(1).find(row => row[0] === clubId);
  if (!clubRow) {
    return jsonOut({ error: 'Club not found' });
  }

  const now = new Date();

  const halls = hallsData
    .filter(h => h[1] === clubId && h[7] === true)
    .map(hall => {
      const hallId = hall[0];
      const comps = computersData.filter(c => c[1] === hallId && c[3] === 'active');

      let occupiedNow = 0;
      comps.forEach(comp => {
        const activeBooking = bookingsData.find(b =>
          b[4] === comp[0] &&
          b[7] === 'confirmed' &&
          new Date(b[5]) <= now &&
          new Date(new Date(b[5]).getTime() + b[6] * 3600000) > now
        );
        if (activeBooking) occupiedNow++;
      });

      const packages = hall[4] ? hall[4].split(',').map(p => p.trim()).filter(p => p) : [];

      return {
        id: hallId,
        name: hall[2],
        base_price_hour: hall[3],
        packages: packages,
        specs: hall[5],
        photo_url: hall[6],
        total_computers: comps.length,
        free_now: comps.length - occupiedNow,
        computers: comps.map(c => ({
          id: c[0],
          number: c[2],
          status: c[3]
        }))
      };
    });

  const club = {
    id: clubRow[0],
    name: clubRow[1],
    address: clubRow[2],
    phone: clubRow[3],
    photo_url: clubRow[4],
    rating: 4.5 + Math.random() * 0.5,
    working_hours: '09:00 – 02:00',
    halls: halls
  };

  return jsonOut(club);
}

// ==================== ИГРОКИ ====================
function registerPlayer(name, phone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const playersSheet = sheet.getSheetByName('players');
  const playersData = playersSheet.getDataRange().getValues();

  const existingPlayer = playersData.slice(1).find(row => row[2] && row[2].toString() === phone);

  if (existingPlayer) {
    return jsonOut({
      id: existingPlayer[0], name: existingPlayer[1], phone: existingPlayer[2],
      bonus_balance: existingPlayer[3], is_new: false
    });
  }

  const newId = playersData.length;
  playersSheet.appendRow([newId, name, phone, 0, new Date().toISOString()]);

  return jsonOut({
    id: newId, name: name, phone: phone, bonus_balance: 0, is_new: true
  });
}

function loginPlayer(phone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const playersData = sheet.getSheetByName('players').getDataRange().getValues();
  const player = playersData.slice(1).find(row => row[2] && row[2].toString() === phone);
  if (player) {
    return jsonOut({
      id: player[0], name: player[1], phone: player[2], bonus_balance: player[3], is_new: false
    });
  }
  return jsonOut({ error: 'Player not found' });
}

// ==================== БРОНИРОВАНИЯ ====================
function createBooking(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const bookingsSheet = sheet.getSheetByName('bookings');
  const playersSheet = sheet.getSheetByName('players');
  const hallsData = sheet.getSheetByName('halls').getDataRange().getValues();
  const computersData = sheet.getSheetByName('computers').getDataRange().getValues();

  const playerPhone = params.player_phone.toString();
  const computerId = parseInt(params.computer_id);
  const startTime = params.start_time.toString();
  const hours = parseInt(params.hours);
  const game = params.game ? params.game.toString() : '';
  const paymentMethod = params.payment_method ? params.payment_method.toString() : 'click';
  const bonusUsed = params.bonus_used ? Number(params.bonus_used) : 0;

  const playersData = playersSheet.getDataRange().getValues();
  const player = playersData.slice(1).find(row => row[2] && row[2].toString() === playerPhone);
  if (!player) {
    return jsonOut({ error: 'Player not found' });
  }

  // Списываем бонусы, если использовались при оплате
  if (bonusUsed > 0) {
    const ok = deductBonusFromPlayer(playerPhone, bonusUsed);
    if (!ok) {
      return jsonOut({ error: 'Недостаточно бонусов для списания' });
    }
  }

  const bookingsData = bookingsSheet.getDataRange().getValues();
  const newId = bookingsData.length;
  bookingsSheet.appendRow([newId, player[1], playerPhone, game, computerId, startTime, hours, 'confirmed', paymentMethod, new Date().toISOString()]);

  // Начисляем бонусы за бронь (% от стоимости часа * часы)
  const computer = computersData.slice(1).find(c => c[0] === computerId);
  const hall = computer ? hallsData.slice(1).find(h => h[0] === computer[1]) : null;
  const pricePerHour = hall ? Number(hall[3]) : 0;
  const total = pricePerHour * hours;
  const bonusEarned = Math.floor(total * BONUS_ACCRUAL_PERCENT / 100);
  addBonusToPlayer(playerPhone, bonusEarned);

  return jsonOut({
    success: true,
    booking_id: newId,
    bonus_earned: bonusEarned,
    new_bonus_balance: getPlayerBonusBalance(playerPhone)
  });
}

function getPlayerBookings(phone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const bookingsData = sheet.getSheetByName('bookings').getDataRange().getValues();
  const computersData = sheet.getSheetByName('computers').getDataRange().getValues();
  const hallsData = sheet.getSheetByName('halls').getDataRange().getValues();
  const clubsData = sheet.getSheetByName('clubs').getDataRange().getValues();

  const now = new Date();

  const playerBookings = bookingsData.slice(1)
    .filter(row => row[2] && row[2].toString() === phone)
    .map(row => {
      const computerId = row[4];
      const startTime = new Date(row[5]);
      const hours = row[6];
      const endTime = new Date(startTime.getTime() + hours * 3600000);

      const computer = computersData.slice(1).find(c => c[0] === computerId);
      const hallId = computer ? computer[1] : null;
      const hall = hallId ? hallsData.slice(1).find(h => h[0] === hallId) : null;
      const clubId = hall ? hall[1] : null;
      const club = clubId ? clubsData.slice(1).find(c => c[0] === clubId) : null;

      let status;
      if (row[7] === 'cancelled') status = 'cancelled';
      else if (endTime < now) status = 'completed';
      else if (startTime <= now && endTime >= now) status = 'active';
      else status = 'upcoming';

      return {
        id: row[0], client_name: row[1], phone: row[2], game: row[3],
        computer_id: computerId, computer_number: computer ? computer[2] : '?',
        start_time: row[5], end_time: endTime.toISOString(), hours: hours,
        status: status, payment_method: row[8],
        club_name: club ? club[1] : 'Неизвестно',
        hall_name: hall ? hall[2] : 'Неизвестно', club_id: clubId
      };
    })
    .sort((a, b) => new Date(b.start_time) - new Date(a.start_time));

  return jsonOut(playerBookings);
}

// ==================== АДМИНКА ====================
function adminLogin(phone, clubId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const clubsData = sheet.getSheetByName('clubs').getDataRange().getValues();
  const club = clubsData.slice(1).find(row => row[0] === clubId);

  if (!club) {
    return jsonOut({ error: 'Club not found' });
  }

  // Пока пускаем всех, кто выбрал клуб (без проверки телефона)
  return jsonOut({
    success: true,
    club_id: clubId,
    club_name: club[1]
  });
}

function getAdminBookings(clubId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const bookingsData = sheet.getSheetByName('bookings').getDataRange().getValues();
  const computersData = sheet.getSheetByName('computers').getDataRange().getValues();
  const hallsData = sheet.getSheetByName('halls').getDataRange().getValues();

  const clubHalls = hallsData.slice(1).filter(h => h[1] === clubId);
  const hallIds = clubHalls.map(h => h[0]);
  const clubComputers = computersData.slice(1).filter(c => hallIds.includes(c[1]));
  const computerIds = clubComputers.map(c => c[0]);

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrow = new Date(today.getTime() + 24 * 3600000);

  const todayBookings = bookingsData.slice(1)
    .filter(b => {
      const startTime = new Date(b[5]);
      return computerIds.includes(b[4]) &&
             b[7] === 'confirmed' &&
             startTime >= today &&
             startTime < tomorrow;
    })
    .map(b => {
      const computer = clubComputers.find(c => c[0] === b[4]);
      const hall = clubHalls.find(h => h[0] === (computer ? computer[1] : null));
      const startTime = new Date(b[5]);
      const endTime = new Date(startTime.getTime() + b[6] * 3600000);

      return {
        id: b[0],
        client_name: b[1],
        phone: b[2],
        game: b[3],
        computer_id: b[4],
        computer_number: computer ? computer[2] : '?',
        hall_name: hall ? hall[2] : '?',
        start_time: b[5],
        end_time: endTime.toISOString(),
        hours: b[6],
        status: b[7]
      };
    });

  const computers = clubComputers.map(c => {
    const hall = clubHalls.find(h => h[0] === c[1]);
    const activeBooking = todayBookings.find(b => {
      const startTime = new Date(b.start_time);
      return b.computer_id === c[0] && startTime <= now && new Date(b.end_time) > now;
    });

    return {
      id: c[0],
      number: c[2],
      hall_id: c[1],
      hall_name: hall ? hall[2] : '?',
      status: c[3],
      is_occupied: !!activeBooking,
      current_booking: activeBooking || null
    };
  });

  return jsonOut({
    club_id: clubId,
    date: today.toISOString(),
    bookings: todayBookings,
    computers: computers
  });
}

function adminCreateBooking(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const bookingsSheet = sheet.getSheetByName('bookings');

  const clientName = params.client_name ? params.client_name.toString() : 'Гость';
  const phone = params.phone ? params.phone.toString() : '';
  const computerId = parseInt(params.computer_id);
  const startTime = params.start_time.toString();
  const hours = parseInt(params.hours);
  const game = params.game ? params.game.toString() : '';

  const bookingsData = bookingsSheet.getDataRange().getValues();
  const newId = bookingsData.length;
  bookingsSheet.appendRow([newId, clientName, phone, game, computerId, startTime, hours, 'confirmed', 'cash', new Date().toISOString()]);

  return jsonOut({ success: true, booking_id: newId });
}

function adminCancelBooking(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const bookingsSheet = sheet.getSheetByName('bookings');
  const bookingId = parseInt(params.booking_id);

  const bookingsData = bookingsSheet.getDataRange().getValues();
  const rowIndex = bookingsData.findIndex(row => row[0] === bookingId);

  if (rowIndex === -1) {
    return jsonOut({ error: 'Booking not found' });
  }

  bookingsSheet.getRange(rowIndex + 1, 8).setValue('cancelled');

  return jsonOut({ success: true });
}

function adminBlockComputer(params) {
  // ИСПРАВЛЕНО: раньше фронт слал параметр "action_type", а бэкенд читал
  // "params.action" — но это имя уже занято роутером (значение всегда было
  // "admin_block_computer"), поэтому block/unblock никогда не срабатывали
  // по назначению и функция просто тумблила статус. Теперь читаем оба
  // варианта имени параметра для совместимости.
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const computersSheet = sheet.getSheetByName('computers');
  const computerId = parseInt(params.computer_id);
  const actionType = (params.action_type || params.block_action || 'toggle').toString();

  const computersData = computersSheet.getDataRange().getValues();
  const rowIndex = computersData.findIndex(row => row[0] === computerId);

  if (rowIndex === -1) {
    return jsonOut({ error: 'Computer not found' });
  }

  const currentStatus = computersData[rowIndex][3];
  let newStatus;

  if (actionType === 'block') newStatus = 'maintenance';
  else if (actionType === 'unblock') newStatus = 'active';
  else newStatus = currentStatus === 'active' ? 'maintenance' : 'active';

  computersSheet.getRange(rowIndex + 1, 4).setValue(newStatus);

  return jsonOut({ success: true, status: newStatus });
}
