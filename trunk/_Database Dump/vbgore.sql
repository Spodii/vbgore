/*
MySQL Data Transfer
Source Host: localhost
Source Database: vbgore
Target Host: localhost
Target Database: vbgore
Date: 3/13/2007 11:44:16 PM
*/

SET FOREIGN_KEY_CHECKS=0;
-- ----------------------------
-- Table structure for banned_ips
-- ----------------------------
CREATE TABLE `banned_ips` (
  `ip` varchar(255) NOT NULL COMMENT 'The IP of the banned address',
  `reason` varchar(255) NOT NULL COMMENT 'The reason the listed IP is banned',
  PRIMARY KEY  (`ip`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for mail
-- ----------------------------
CREATE TABLE `mail` (
  `id` int(11) NOT NULL COMMENT 'ID of the mail',
  `sub` varchar(255) NOT NULL COMMENT 'Subject text',
  `by` varchar(255) NOT NULL COMMENT 'Mail writer name',
  `date` date NOT NULL COMMENT 'Date the mail was recieved',
  `msg` text NOT NULL COMMENT 'Body message',
  `new` tinyint(3) unsigned NOT NULL default '0' COMMENT 'If the mail is new (1 = yes, 0 = no)',
  `objs` mediumtext NOT NULL COMMENT 'Objects contained in message (obj index and amount)',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for mail_lastid
-- ----------------------------
CREATE TABLE `mail_lastid` (
  `lastid` int(11) NOT NULL default '0',
  PRIMARY KEY  (`lastid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for npcs
-- ----------------------------
CREATE TABLE `npcs` (
  `id` smallint(6) NOT NULL default '0' COMMENT 'Identifier of the NPC',
  `name` varchar(255) NOT NULL COMMENT 'Name',
  `descr` varchar(255) NOT NULL COMMENT 'Description',
  `ai` tinyint(3) unsigned NOT NULL default '0' COMMENT 'AI algorithm used (see Server.NPC.NPC_AI)',
  `chat` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Index of the NPC chat from the NPC Chat file',
  `respawnwait` int(11) NOT NULL default '0' COMMENT 'Time it takes to respawn (in miliseconds)',
  `attackable` tinyint(3) unsigned NOT NULL default '0' COMMENT 'If the NPC is attackable (1 = yes, 0 = no)',
  `attackgrh` int(11) NOT NULL default '0' COMMENT 'Grh the NPC uses when attacking (works like UseGrh)',
  `attackrange` tinyint(3) unsigned NOT NULL default '0' COMMENT 'If the NPC has a ranged attack (0 or 1 for melee)',
  `attacksfx` tinyint(3) unsigned NOT NULL,
  `projectilerotatespeed` tinyint(3) unsigned NOT NULL default '0' COMMENT 'If a ranged attack, how fast the projectile rotates',
  `hostile` tinyint(3) unsigned NOT NULL default '0' COMMENT 'If the NPC is hostile (1 = yes, 0 = no)',
  `quest` smallint(6) NOT NULL default '0' COMMENT 'ID of the quest the NPC gives',
  `drops` mediumtext NOT NULL COMMENT 'List of NPC drops',
  `give_exp` int(11) NOT NULL default '0' COMMENT 'Experience given upon killing the NPC',
  `give_gold` int(11) NOT NULL default '0' COMMENT 'Gold given upon killing the NPC',
  `objs_shop` mediumtext NOT NULL COMMENT 'Objects sold as a shopkeeper/vendor',
  `char_hair` smallint(6) NOT NULL default '1' COMMENT 'Paperdolling hair ID',
  `char_head` smallint(6) NOT NULL default '1' COMMENT 'Paperdolling head ID',
  `char_body` smallint(6) NOT NULL default '1' COMMENT 'Paperdolling body ID',
  `char_weapon` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling weapon ID',
  `char_wings` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling wings ID',
  `char_heading` tinyint(3) unsigned NOT NULL default '3' COMMENT 'Starting heading (direction the body/etc faces)',
  `char_headheading` tinyint(3) unsigned NOT NULL default '3' COMMENT 'Starting head heading (direction the head faces)',
  `stat_mag` int(11) NOT NULL default '0' COMMENT 'Magic',
  `stat_def` int(11) NOT NULL default '0' COMMENT 'Defence',
  `stat_speed` int(11) NOT NULL default '0' COMMENT 'Walk speed',
  `stat_hit_min` int(11) NOT NULL default '1' COMMENT 'Minimum hit',
  `stat_hit_max` int(11) NOT NULL default '1' COMMENT 'Maximum hit',
  `stat_hp` int(11) NOT NULL default '10' COMMENT 'Health points',
  `stat_mp` int(11) NOT NULL default '10' COMMENT 'Mana points',
  `stat_sp` int(11) NOT NULL default '10' COMMENT 'Stamina points',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for objects
-- ----------------------------
CREATE TABLE `objects` (
  `id` smallint(6) NOT NULL COMMENT 'Identifier of the object',
  `name` varchar(255) NOT NULL COMMENT 'Name',
  `price` int(11) NOT NULL default '0' COMMENT 'Price object is bought for',
  `objtype` tinyint(3) unsigned NOT NULL COMMENT 'Object type (see Server.Declares for OBJTYPE_ consts)',
  `weapontype` tinyint(3) unsigned NOT NULL COMMENT 'Weapon type (Only valid if obj=weapon - see Server.Declares)',
  `weaponrange` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Range of the weapon''s attack (if ranged)',
  `grhindex` int(11) NOT NULL COMMENT 'Index of the object graphic (by Grh value)',
  `usegrh` int(11) NOT NULL default '0' COMMENT 'Grh for the weapon''s attack',
  `usesfx` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Sound played when the object is used (0 for none)',
  `projectilerotatespeed` tinyint(3) unsigned NOT NULL COMMENT 'If a projectile, how fast it rotates (0 for no rotate)',
  `sprite_body` smallint(6) NOT NULL default '-1' COMMENT 'Paperdolling body changed to upon usage (-1 for no change)',
  `sprite_weapon` smallint(6) NOT NULL default '-1' COMMENT 'Paperdolling weapon changed to upon usage (-1 for no change)',
  `sprite_hair` smallint(6) NOT NULL default '-1' COMMENT 'Paperdolling hair changed to upon usage (-1 for no change)',
  `sprite_head` smallint(6) NOT NULL default '-1' COMMENT 'Paperdolling head changed to upon usage (-1 for no change)',
  `sprite_wings` smallint(6) NOT NULL default '-1' COMMENT 'Paperdolling wings changed to upon usage (-1 for no change)',
  `replenish_hp` int(11) NOT NULL default '0' COMMENT 'Amount of HP replenished upon usage',
  `replenish_mp` int(11) NOT NULL default '0' COMMENT 'Amount of MP replenished upon usage',
  `replenish_sp` int(11) NOT NULL default '0' COMMENT 'Amount of SP replenished upon usage',
  `replenish_hp_percent` int(11) NOT NULL default '0' COMMENT 'Percent of HP replenished upon usage',
  `replenish_mp_percent` int(11) NOT NULL default '0' COMMENT 'Percent of MP replenished upon usage',
  `replenish_sp_percent` int(11) NOT NULL default '0' COMMENT 'Percent of SP replenished upon usage',
  `stat_str` int(11) NOT NULL default '0' COMMENT 'Strength raised upon usage',
  `stat_agi` int(11) NOT NULL default '0' COMMENT 'Agility raised upon usage',
  `stat_mag` int(11) NOT NULL default '0' COMMENT 'Magic raised upon usage',
  `stat_def` int(11) NOT NULL default '0' COMMENT 'Defence raised upon usage',
  `stat_speed` int(11) NOT NULL default '0' COMMENT 'Walk speed raised upon usage',
  `stat_hit_min` int(11) NOT NULL default '0' COMMENT 'Minimum hit raised upon usage',
  `stat_hit_max` int(11) NOT NULL default '0' COMMENT 'Maximum hit raised upon usage',
  `stat_hp` int(11) NOT NULL default '0' COMMENT 'Health raised upon usage',
  `stat_mp` int(11) NOT NULL default '0' COMMENT 'Magic raised upon usage',
  `stat_sp` int(11) NOT NULL default '0' COMMENT 'Stamina raised upon usage',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for quests
-- ----------------------------
CREATE TABLE `quests` (
  `id` smallint(6) NOT NULL COMMENT 'Identifier of the quest',
  `name` varchar(255) NOT NULL COMMENT 'Name',
  `redoable` tinyint(3) unsigned NOT NULL default '0' COMMENT 'If the quest is redoable (1 = yes, 0 = no)',
  `text_start` varchar(255) NOT NULL COMMENT 'Text said at start of quest',
  `text_accept` varchar(255) NOT NULL COMMENT 'Text said when accepting a quest',
  `text_incomplete` varchar(255) NOT NULL COMMENT 'Text said when trying to finish a quest (reqs. not met)',
  `text_finish` varchar(255) NOT NULL COMMENT 'Text said when finishing a quest (requirements met)',
  `text_info` text NOT NULL COMMENT 'All the quest info that appears on the client quest screen',
  `accept_req_level` int(11) NOT NULL default '0' COMMENT 'Required level to accept',
  `accept_req_obj` smallint(6) NOT NULL default '0' COMMENT 'Required object to accept (object ID)',
  `accept_req_objamount` smallint(6) NOT NULL default '0' COMMENT 'Required object amount to accept (if accept_req_obj > 0)',
  `accept_reward_exp` int(11) NOT NULL default '0' COMMENT 'Experience reward upon accepting',
  `accept_reward_gold` int(11) NOT NULL default '0' COMMENT 'Gold reward upon accepting',
  `accept_reward_obj` smallint(6) NOT NULL default '0' COMMENT 'Object reward upon accepting',
  `accept_reward_objamount` smallint(6) NOT NULL default '0' COMMENT 'Object amount reward upon accepting (accept_reward_obj > 0)',
  `accept_reward_learnskill` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Skill learned upon accepting',
  `finish_req_obj` smallint(6) NOT NULL default '0' COMMENT 'Required object to finish (object ID)',
  `finish_req_objamount` smallint(6) NOT NULL default '0' COMMENT 'Required object amount to finish (if finish_req_obj > 0)',
  `finish_req_killnpc` smallint(6) NOT NULL default '0' COMMENT 'Index of the NPC to kill to complete quest',
  `finish_req_killnpcamount` smallint(6) NOT NULL default '0' COMMENT 'Number of the NPCs to kill (if killnpc > 0) to finish quest',
  `finish_reward_exp` int(11) NOT NULL default '0' COMMENT 'Experience reward upon finishing',
  `finish_reward_gold` int(11) NOT NULL default '0' COMMENT 'Gold reward upon finishing',
  `finish_reward_obj` smallint(6) NOT NULL default '0' COMMENT 'Object reward upon finishing',
  `finish_reward_objamount` smallint(6) NOT NULL default '0' COMMENT 'Object amount reward upon finishing (finish_reward_obj > 0)',
  `finish_reward_learnskill` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Skill learned upon finishing',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for users
-- ----------------------------
CREATE TABLE `users` (
  `name` varchar(255) NOT NULL COMMENT 'Name',
  `ip` varchar(255) NOT NULL COMMENT 'Holds the last 10 IPs used to connect to this account',
  `gm` tinyint(3) unsigned NOT NULL default '0',
  `password` varchar(255) NOT NULL COMMENT 'Password',
  `descr` varchar(255) NOT NULL COMMENT 'Description',
  `inventory` text NOT NULL,
  `bank` text NOT NULL COMMENT 'List of the user''s bank items',
  `mail` text NOT NULL,
  `knownskills` text NOT NULL COMMENT 'Skills known by the user (1 = known, 0 = unknown)',
  `completedquests` text NOT NULL COMMENT 'Defines the quests completed (not recommended to edit)',
  `currentquest` text NOT NULL COMMENT 'Quest(s) the user is currently on (do not edit)',
  `date_create` date NOT NULL COMMENT 'The date the account was created',
  `date_lastlogin` date NOT NULL COMMENT 'The date the user last logged in',
  `pos_x` tinyint(3) unsigned NOT NULL COMMENT 'X position',
  `pos_y` tinyint(3) unsigned NOT NULL COMMENT 'Y position',
  `pos_map` smallint(6) NOT NULL COMMENT 'Map',
  `char_hair` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling hair',
  `char_head` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling head',
  `char_body` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling body',
  `char_weapon` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling weapon',
  `char_wings` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling wings',
  `char_heading` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Direction the character is pointed towards',
  `char_headheading` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Direction the character''s head is pointed towards',
  `eq_weapon` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Slot of equipted weapon',
  `eq_armor` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Slot of equipted armor',
  `eq_wings` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Slot of equipted wings',
  `stat_str` int(11) NOT NULL default '0' COMMENT 'Base strength',
  `stat_agi` int(11) NOT NULL default '0' COMMENT 'Base agility',
  `stat_mag` int(11) NOT NULL default '0' COMMENT 'Base magic',
  `stat_def` int(11) NOT NULL default '0' COMMENT 'Base defense',
  `stat_speed` int(11) NOT NULL COMMENT 'Base walking speed',
  `stat_gold` int(11) NOT NULL default '0' COMMENT 'Gold',
  `stat_exp` int(11) NOT NULL default '0' COMMENT 'Experience',
  `stat_elv` int(11) NOT NULL default '0' COMMENT 'Level',
  `stat_elu` int(11) NOT NULL default '0' COMMENT 'Experience required for next level',
  `stat_points` int(11) NOT NULL default '0' COMMENT 'Points in update queue',
  `stat_hit_min` int(11) NOT NULL default '0' COMMENT 'Base minimum hit damage',
  `stat_hit_max` int(11) NOT NULL default '0' COMMENT 'Base maximum hit damage',
  `stat_hp_min` int(11) NOT NULL default '0' COMMENT 'Current health',
  `stat_hp_max` int(11) NOT NULL default '0' COMMENT 'Base maximum health',
  `stat_mp_min` int(11) NOT NULL default '0' COMMENT 'Current mana',
  `stat_mp_max` int(11) NOT NULL default '0' COMMENT 'Base maximum mana',
  `stat_sp_min` int(11) NOT NULL default '0' COMMENT 'Current stamina',
  `stat_sp_max` int(11) NOT NULL default '0' COMMENT 'Base maximum stamina',
  `server` tinyint(4) unsigned NOT NULL default '0' COMMENT 'States what server the user is on (0 = not online)',
  PRIMARY KEY  (`name`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records 
-- ----------------------------
INSERT INTO `mail` VALUES ('1', 'Test Message', 'Game Admin', '2007-03-13', 'This is a test message that simply shows the pwnification of the mailing system. Here, have a random number! 4.535275', '1', '5 6\r\n5 3\r\n3 8\r\n1 8\r\n6 8');
INSERT INTO `mail` VALUES ('2', 'Test Message', 'Game Admin', '2007-03-13', 'This is a test message that simply shows the pwnification of the mailing system. Here, have a random number! 41.40327', '1', '5 6\r\n5 3\r\n3 8\r\n1 8\r\n6 8');
INSERT INTO `mail` VALUES ('3', 'Test Message', 'Game Admin', '2007-03-13', 'This is a test message that simply shows the pwnification of the mailing system. Here, have a random number! 86.26193', '1', '5 6\r\n5 3\r\n3 8\r\n1 8\r\n6 8');
INSERT INTO `mail` VALUES ('4', 'Test Message', 'Game Admin', '2007-03-13', 'This is a test message that simply shows the pwnification of the mailing system. Here, have a random number! 79.048', '1', '5 6\r\n5 3\r\n3 8\r\n1 8\r\n6 8');
INSERT INTO `mail` VALUES ('5', 'Test Message', 'Game Admin', '2007-03-13', 'This is a test message that simply shows the pwnification of the mailing system. Here, have a random number! 37.35362', '1', '5 6\r\n5 3\r\n3 8\r\n1 8\r\n6 8');
INSERT INTO `mail_lastid` VALUES ('5');
INSERT INTO `npcs` VALUES ('1', 'Headless Man', 'This man seems to want your help!', '0', '3', '0', '0', '0', '0', '0', '0', '0', '1', '', '0', '0', '', '1', '0', '1', '0', '1', '3', '3', '0', '0', '3', '1', '1', '10', '10', '10');
INSERT INTO `npcs` VALUES ('2', 'Bandit', 'Bald little rascal who wants your booty!', '3', '2', '5000', '1', '26', '0', '1', '0', '1', '0', '1 2 50\r\n5 1 10\r\n6 1 10\r\n7 1 10', '10', '10', '', '0', '1', '1', '1', '0', '3', '3', '0', '0', '3', '1', '2', '15', '2', '2');
INSERT INTO `npcs` VALUES ('3', 'Shopkeeper', 'Just a humble shopkeeper.', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '', '0', '0', '1 -1\r\n2 -1\r\n3 -1\r\n4 -1\r\n5 -1\r\n6 -1\r\n7 -1', '1', '1', '1', '0', '1', '3', '3', '0', '0', '3', '1', '1', '10', '10', '10');
INSERT INTO `npcs` VALUES ('4', 'Ninja', 'A sneaky little ninja with a hand full of ninja stars', '4', '2', '10000', '1', '11', '10', '7', '100', '1', '0', '1 2 50\r\n5 1 10\r\n6 1 10\r\n7 1 10', '25', '20', '', '0', '1', '1', '1', '1', '3', '3', '0', '0', '5', '2', '4', '10', '10', '10');
INSERT INTO `npcs` VALUES ('5', 'Cleric', 'Holy practicer of the church\'s arts', '5', '2', '15000', '1', '26', '0', '0', '0', '1', '0', '1 2 50\r\n5 1 10\r\n6 1 10\r\n7 1 10', '50', '50', '', '1', '1', '1', '0', '1', '3', '3', '1', '0', '3', '1', '1', '10', '50', '10');
INSERT INTO `npcs` VALUES ('6', 'Banker', 'A wealthy little bank owner', '6', '0', '0', '0', '0', '0', '0', '0', '0', '0', '', '0', '0', '', '1', '1', '1', '0', '1', '3', '3', '0', '0', '0', '1', '1', '10', '10', '10');
INSERT INTO `npcs` VALUES ('7', 'Crazy man', 'Crazy man rambling about everything and nothing', '2', '1', '0', '0', '0', '0', '0', '0', '0', '0', '', '0', '0', '', '1', '1', '1', '0', '0', '3', '3', '0', '0', '0', '1', '1', '10', '10', '10');
INSERT INTO `objects` VALUES ('1', 'Tiny Healing Potion', '10', '1', '0', '0', '38', '0', '0', '0', '-1', '-1', '-1', '-1', '-1', '10', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `objects` VALUES ('2', 'Mini Healing Potion', '10', '1', '0', '0', '38', '0', '0', '0', '-1', '-1', '-1', '-1', '-1', '20', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `objects` VALUES ('3', 'Small Healing Potion', '10', '1', '0', '0', '38', '0', '0', '0', '-1', '-1', '-1', '-1', '-1', '30', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `objects` VALUES ('4', 'Healing Potion', '10', '1', '0', '0', '38', '0', '0', '0', '-1', '-1', '-1', '-1', '-1', '100', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `objects` VALUES ('5', 'Newbie Armor', '10', '3', '0', '0', '1000', '0', '0', '0', '2', '-1', '-1', '-1', '-1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '3', '0', '0', '0', '0', '0', '0');
INSERT INTO `objects` VALUES ('6', 'Newbie Dagger', '30', '2', '1', '0', '1300', '26', '1', '0', '-1', '1', '-1', '-1', '-1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '2', '4', '0', '0', '0');
INSERT INTO `objects` VALUES ('7', 'Angel Wings', '100', '4', '0', '0', '1200', '0', '0', '0', '-1', '-1', '-1', '-1', '1', '0', '0', '0', '0', '0', '0', '1', '1', '1', '1', '0', '1', '1', '20', '10', '10');
INSERT INTO `objects` VALUES ('8', 'Ninja Stars', '100', '2', '4', '10', '11', '11', '7', '100', '-1', '0', '-1', '-1', '-1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '1', '6', '0', '0', '0');
INSERT INTO `objects` VALUES ('9', 'Big Star', '15', '1', '0', '0', '27', '14', '0', '0', '-1', '-1', '-1', '-1', '-1', '0', '0', '0', '0', '100', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `quests` VALUES ('1', 'Kill Bandits', '1', 'Help me get revenge!', 'Thanks for the help! Kill 3 bandits that hide in the waterfall!', 'Just because I have no head doesn\'t mean I have no brain...', 'Sweet d00d, that\'ll show them whos boss! ^_^', 'The Headless Man has told you about some dangerous bandits that have nested in the cave under the |waterfall| in the west side of the island, outside of town. They have been stealing junk from the only two houses on this pathetic island, and it is important that we get it back, since without our junk, we are useless.\r\n\r\nTalk to the Headless Man after you kill the 3 bandits for your reward.', '1', '0', '0', '50', '45', '0', '0', '1', '0', '0', '2', '3', '200', '400', '2', '60', '2');
INSERT INTO `users` VALUES ('Spodi', '127.0.0.1', '1', 'f887eb538bb69342ac792536bcdaf02d', '', '1 1 5 0\r\n2 2 1 0\r\n3 3 1 0\r\n4 5 1 1\r\n5 6 1 1\r\n6 7 1 1\r\n7 8 1 0\r\n8 9 50 0', '', '1\r\n2\r\n3\r\n4\r\n5', '1\r\n2\r\n3\r\n4\r\n5\r\n6\r\n7', '', '', '2007-03-13', '2007-03-13', '42', '34', '1', '1', '1', '2', '1', '1', '3', '3', '5', '4', '6', '1', '1', '1', '1', '5', '100', '0', '1', '10', '0', '1', '1', '60', '50', '60', '50', '4', '50', '0');
