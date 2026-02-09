const crypto = require('crypto');
const id = 'AKfycbxYvGtsYldsjhvzLTxxevy7Fm0Jhb2VNqzY1FXZayjGaBjx0WsqjFj2cmuAuPC2wmGU';
const hash = crypto.createHash('sha256').update(id).digest('hex');
console.log(hash);
