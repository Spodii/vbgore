/*
MySQL Data Transfer
Source Host: localhost
Source Database: vbgore
Target Host: localhost
Target Database: vbgore
Date: 12/7/2006 12:01:00 AM
*/

SET FOREIGN_KEY_CHECKS=0;
-- ----------------------------
-- Table structure for mail
-- ----------------------------
CREATE TABLE `mail` (
  `id` int(11) NOT NULL COMMENT 'ID of the mail',
  `sub` varchar(255) NOT NULL COMMENT 'Subject text',
  `by` varchar(255) NOT NULL COMMENT 'Mail writer name',
  `date` date NOT NULL COMMENT 'Date the mail was recieved',
  `msg` text NOT NULL COMMENT 'Body message',
  `new` tinyint(4) NOT NULL default '0' COMMENT 'If the mail is new (1 = yes, 0 = no)',
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
  `desc` varchar(255) NOT NULL COMMENT 'Description',
  `movement` smallint(6) default '0' COMMENT 'Movement style (see Server.NPC.NPC_AI)',
  `respawnwait` int(11) default '0' COMMENT 'Time it takes to respawn (in miliseconds)',
  `attackable` tinyint(4) default '0' COMMENT 'If the NPC is attackable (1 = yes, 0 = no)',
  `hostile` tinyint(4) default '0' COMMENT 'If the NPC is hostile (1 = yes, 0 = no)',
  `quest` smallint(6) default '0' COMMENT 'ID of the quest the NPC gives',
  `give_exp` int(11) default '0' COMMENT 'Experience given upon killing the NPC',
  `give_gold` int(11) default '0' COMMENT 'Gold given upon killing the NPC',
  `objs_shop` mediumtext COMMENT 'Objects sold as a shopkeeper/vendor',
  `char_hair` smallint(6) default '1' COMMENT 'Paperdolling hair ID',
  `char_head` smallint(6) default '1' COMMENT 'Paperdolling head ID',
  `char_body` smallint(6) default '1' COMMENT 'Paperdolling body ID',
  `char_weapon` smallint(6) default '0' COMMENT 'Paperdolling weapon ID',
  `char_wings` smallint(6) default '0' COMMENT 'Paperdolling wings ID',
  `char_heading` tinyint(4) default '3' COMMENT 'Starting heading (direction the body/etc faces)',
  `char_headheading` tinyint(4) default '3' COMMENT 'Starting head heading (direction the head faces)',
  `stat_mag` int(11) default '0' COMMENT 'Magic',
  `stat_def` int(11) default '0' COMMENT 'Defence',
  `stat_hit_min` int(11) default '1' COMMENT 'Minimum hit',
  `stat_hit_max` int(11) default '1' COMMENT 'Maximum hit',
  `stat_hp` int(11) default '10' COMMENT 'Health points',
  `stat_mp` int(11) default '10' COMMENT 'Mana points',
  `stat_sp` int(11) default '10' COMMENT 'Stamina points',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for objects
-- ----------------------------
CREATE TABLE `objects` (
  `id` smallint(6) NOT NULL COMMENT 'Identifier of the object',
  `name` varchar(255) NOT NULL COMMENT 'Name',
  `price` int(11) NOT NULL default '0' COMMENT 'Price object is bought for',
  `objtype` tinyint(4) NOT NULL COMMENT 'Object type (see Server.Declares for OBJTYPE_ consts)',
  `weapontype` tinyint(4) default NULL COMMENT 'Weapon type (Only valid if obj=weapon - see Server.Declares)',
  `grhindex` int(11) NOT NULL COMMENT 'Index of the object graphic (by Grh value)',
  `sprite_body` smallint(6) default '-1' COMMENT 'Paperdolling body changed to upon usage (-1 for no change)',
  `sprite_weapon` smallint(6) default '-1' COMMENT 'Paperdolling weapon changed to upon usage (-1 for no change)',
  `sprite_hair` smallint(6) default '-1' COMMENT 'Paperdolling hair changed to upon usage (-1 for no change)',
  `sprite_head` smallint(6) default '-1' COMMENT 'Paperdolling head changed to upon usage (-1 for no change)',
  `sprite_wings` smallint(6) default '-1' COMMENT 'Paperdolling wings changed to upon usage (-1 for no change)',
  `replenish_hp` int(11) default '0' COMMENT 'Amount of HP replenished upon usage',
  `replenish_mp` int(11) default '0' COMMENT 'Amount of MP replenished upon usage',
  `replenish_sp` int(11) default '0' COMMENT 'Amount of SP replenished upon usage',
  `replenish_hp_percent` int(11) default '0' COMMENT 'Percent of HP replenished upon usage',
  `replenish_mp_percent` int(11) default '0' COMMENT 'Percent of MP replenished upon usage',
  `replenish_sp_percent` int(11) default '0' COMMENT 'Percent of SP replenished upon usage',
  `stat_str` int(11) default '0' COMMENT 'Strength raised upon usage',
  `stat_agi` int(11) default '0' COMMENT 'Agility raised upon usage',
  `stat_mag` int(11) default '0' COMMENT 'Magic raised upon usage',
  `stat_def` int(11) default '0' COMMENT 'Defence raised upon usage',
  `stat_hit_min` int(11) default '0' COMMENT 'Minimum hit raised upon usage',
  `stat_hit_max` int(11) default '0' COMMENT 'Maximum hit raised upon usage',
  `stat_hp` int(11) default '0' COMMENT 'Health raised upon usage',
  `stat_mp` int(11) default '0' COMMENT 'Magic raised upon usage',
  `stat_sp` int(11) default '0' COMMENT 'Stamina raised upon usage',
  `stat_exp` int(11) default '0' COMMENT 'Experienced raised upon usage',
  `stat_points` int(11) default '0' COMMENT 'Update points raised upon usage',
  `stat_gold` int(11) default '0' COMMENT 'Gold raised upon usage',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for quests
-- ----------------------------
CREATE TABLE `quests` (
  `id` smallint(6) NOT NULL COMMENT 'Identifier of the quest',
  `name` varchar(255) NOT NULL COMMENT 'Name',
  `redoable` tinyint(4) NOT NULL default '0' COMMENT 'If the quest is redoable (1 = yes, 0 = no)',
  `text_start` varchar(255) NOT NULL COMMENT 'Text said at start of quest',
  `text_accept` varchar(255) NOT NULL COMMENT 'Text said when accepting a quest',
  `text_incomplete` varchar(255) NOT NULL COMMENT 'Text said when trying to finish a quest (reqs. not met)',
  `text_finish` varchar(255) NOT NULL COMMENT 'Text said when finishing a quest (requirements met)',
  `accept_req_level` int(11) default '0' COMMENT 'Required level to accept',
  `accept_req_obj` smallint(6) default '0' COMMENT 'Required object to accept (object ID)',
  `accept_req_objamount` smallint(6) default '0' COMMENT 'Required object amount to accept (if accept_req_obj > 0)',
  `accept_reward_exp` int(11) default '0' COMMENT 'Experience reward upon accepting',
  `accept_reward_gold` int(11) default '0' COMMENT 'Gold reward upon accepting',
  `accept_reward_obj` smallint(6) default '0' COMMENT 'Object reward upon accepting',
  `accept_reward_objamount` smallint(6) default '0' COMMENT 'Object amount reward upon accepting (accept_reward_obj > 0)',
  `accept_reward_learnskill` tinyint(4) default '0' COMMENT 'Skill learned upon accepting',
  `finish_req_obj` smallint(6) default '0' COMMENT 'Required object to finish (object ID)',
  `finish_req_objamount` smallint(6) default '0' COMMENT 'Required object amount to finish (if finish_req_obj > 0)',
  `finish_req_killnpc` smallint(6) default '0' COMMENT 'Index of the NPC to kill to complete quest',
  `finish_req_killnpcamount` smallint(6) default '0' COMMENT 'Number of the NPCs to kill (if killnpc > 0) to finish quest',
  `finish_reward_exp` int(11) default '0' COMMENT 'Experience reward upon finishing',
  `finish_reward_gold` int(11) default '0' COMMENT 'Gold reward upon finishing',
  `finish_reward_obj` smallint(6) default '0' COMMENT 'Object reward upon finishing',
  `finish_reward_objamount` smallint(6) default '0' COMMENT 'Object amount reward upon finishing (finish_reward_obj > 0)',
  `finish_reward_learnskill` tinyint(4) default '0' COMMENT 'Skill learned upon finishing',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for users
-- ----------------------------
CREATE TABLE `users` (
  `name` varchar(255) NOT NULL COMMENT 'Name',
  `gm` tinyint(4) NOT NULL,
  `password` varchar(255) NOT NULL COMMENT 'Password',
  `desc` varchar(255) NOT NULL COMMENT 'Description',
  `inventory` mediumtext NOT NULL,
  `mail` mediumtext NOT NULL,
  `knownskills` text NOT NULL COMMENT 'Skills known by the user (1 = known, 0 = unknown)',
  `completedquests` mediumtext NOT NULL COMMENT 'Defines the quests completed (not recommended to edit)',
  `currentquest` mediumtext NOT NULL COMMENT 'Quest(s) the user is currently on (do not edit)',
  `pos_x` tinyint(4) NOT NULL COMMENT 'X position',
  `pos_y` tinyint(4) NOT NULL COMMENT 'Y position',
  `pos_map` smallint(6) NOT NULL COMMENT 'Map',
  `char_hair` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling hair',
  `char_head` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling head',
  `char_body` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling body',
  `char_weapon` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling weapon',
  `char_wings` smallint(6) NOT NULL default '0' COMMENT 'Paperdolling wings',
  `char_heading` tinyint(4) NOT NULL default '0' COMMENT 'Direction the character is pointed towards',
  `char_headheading` tinyint(4) NOT NULL default '0' COMMENT 'Direction the character''s head is pointed towards',
  `eq_weapon` tinyint(4) NOT NULL default '0' COMMENT 'Slot of equipted weapon',
  `eq_armor` tinyint(4) NOT NULL default '0' COMMENT 'Slot of equipted armor',
  `eq_wings` tinyint(4) NOT NULL default '0' COMMENT 'Slot of equipted wings',
  `stat_str` int(11) NOT NULL default '0' COMMENT 'Base strength',
  `stat_agi` int(11) NOT NULL default '0' COMMENT 'Base agility',
  `stat_mag` int(11) NOT NULL default '0' COMMENT 'Base magic',
  `stat_def` int(11) NOT NULL default '0' COMMENT 'Base defense',
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
  `online` tinyint(4) NOT NULL default '0' COMMENT 'States if the user is online or not (1 for yes, 0 for no)',
  PRIMARY KEY  (`name`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records 
-- ----------------------------
INSERT INTO `mail_lastid` VALUES ('0');
INSERT INTO `npcs` VALUES ('1', 'Headless Man', 'This man seems to want your help!', '0', '0', '0', '0', '1', '0', '0', '', '1', '0', '1', '0', '1', '3', '3', '0', '0', '1', '1', '10', '10', '10');
INSERT INTO `npcs` VALUES ('2', 'Bandit', 'Bald little rascal who wants your booty!', '3', '5000', '1', '1', '0', '10', '10', '', '0', '1', '1', '1', '0', '3', '3', '0', '0', '1', '2', '2', '2', '2');
INSERT INTO `npcs` VALUES ('3', 'Shopkeeper', 'Just a humble shopkeeper.', '0', '0', '0', '0', '0', '0', '0', '1 -1\r\n2 -1\r\n3 -1\r\n4 -1\r\n5 -1\r\n6 -1\r\n7 -1', '1', '1', '1', '0', '1', '3', '3', '0', '0', '1', '1', '10', '10', '10');
INSERT INTO `objects` VALUES ('1', 'Healing Potion', '10', '1', '0', '4', '-1', '-1', '-1', '-1', '-1', '100', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `objects` VALUES ('2', 'Healing Potion', '10', '1', '0', '4', '-1', '-1', '-1', '-1', '-1', '100', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `objects` VALUES ('3', 'Healing Potion', '10', '1', '0', '4', '-1', '-1', '-1', '-1', '-1', '100', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `objects` VALUES ('4', 'Healing Potion', '10', '1', '0', '4', '-1', '-1', '-1', '-1', '-1', '100', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `objects` VALUES ('5', 'Newbie Armor', '10', '3', '0', '1000', '2', '-1', '-1', '-1', '-1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '3', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `objects` VALUES ('6', 'Newbie Dagger', '30', '2', '1', '1300', '-1', '1', '-1', '-1', '-1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '2', '4', '0', '0', '0', '0', '0', '0');
INSERT INTO `objects` VALUES ('7', 'Angel Wings', '100', '4', '0', '1200', '-1', '-1', '-1', '-1', '1', '0', '0', '0', '0', '0', '0', '1', '1', '1', '1', '1', '1', '20', '10', '10', '0', '0', '0');
INSERT INTO `quests` VALUES ('1', 'Kill Bandits', '1', 'Help me get revenge!', 'Thanks for the help! Kill 3 bandits that hide in the waterfall!', 'Just because I have no head doesn\'t mean I have no brain...', 'Sweet d00d, that\'ll show them whos boss! ^_^', '1', '0', '0', '100', '0', '0', '0', '8', '0', '0', '2', '3', '200', '400', '2', '60', '1');
