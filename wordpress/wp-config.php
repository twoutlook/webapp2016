<?php
/**
 * The base configuration for WordPress
 *
 * The wp-config.php creation script uses this file during the
 * installation. You don't have to use the web site, you can
 * copy this file to "wp-config.php" and fill in the values.
 *
 * This file contains the following configurations:
 *
 * * MySQL settings
 * * Secret keys
 * * Database table prefix
 * * ABSPATH
 *
 * @link https://codex.wordpress.org/Editing_wp-config.php
 *
 * @package WordPress
 */

// ** MySQL settings - You can get this info from your web host ** //
/** The name of the database for WordPress */
define('DB_NAME', 'wp453');

/** MySQL database username */
define('DB_USER', 'wp453');

/** MySQL database password */
define('DB_PASSWORD', 'hJGGTshJXXjBHb9v');

/** MySQL hostname */
define('DB_HOST', 'localhost');

/** Database Charset to use in creating database tables. */
define('DB_CHARSET', 'utf8mb4');

/** The Database Collate type. Don't change this if in doubt. */
define('DB_COLLATE', '');

/**#@+
 * Authentication Unique Keys and Salts.
 *
 * Change these to different unique phrases!
 * You can generate these using the {@link https://api.wordpress.org/secret-key/1.1/salt/ WordPress.org secret-key service}
 * You can change these at any point in time to invalidate all existing cookies. This will force all users to have to log in again.
 *
 * @since 2.6.0
 */
define('AUTH_KEY',         ':`-@)-~)RVljY]3hwxQT6Lx<l%lUIH}Y&MzV<BU{:B)&6Jy6Uj,#o#j<+2/4g05G');
define('SECURE_AUTH_KEY',  'N~XL#pH <3(]b1)J|I~b;hDgx_SpZKTMZs!cft;Pgi^9F9~6Y$xOes0Mo+/G/tEN');
define('LOGGED_IN_KEY',    'x[_Za2v,aO=BKMbk&%<,b.xN*do:#&4F(fBvMfh88iLe&4X$op121.CN4x5d7KG0');
define('NONCE_KEY',        ' M6+com h4Z>zm|ArA{X&P4fRTa?g`Hi>YQa:,::c%)b1C`(M)OW#l+?l>b:Xj2]');
define('AUTH_SALT',        '9Ch=Pgs:=hkQ}d8})>7.j%{TS]U~[8x!qGHI`]Bqo3So-xM`/qW1#upn>@GYOCOh');
define('SECURE_AUTH_SALT', '=BXv >0n?QMUk`DIM7@]KJm53YUYGl%dYE&K*9hQg)Pa>a;F-:J,%*4&M{E?&oLJ');
define('LOGGED_IN_SALT',   ' +KaXu]^!CoAFg2y.uC5qCJ0%|wl]BhN$Y1eLD8-;?e)?]YxQJv<@-k%9e&H5w8/');
define('NONCE_SALT',       '{>B({:Hwosn<dt4KtzRHkv3,Q_/#{|*,G&nio?CFny,JSP=A@fHHT&<lHL!V9Bc}');

/**#@-*/

/**
 * WordPress Database Table prefix.
 *
 * You can have multiple installations in one database if you give each
 * a unique prefix. Only numbers, letters, and underscores please!
 */
$table_prefix  = 'wp_';

/**
 * For developers: WordPress debugging mode.
 *
 * Change this to true to enable the display of notices during development.
 * It is strongly recommended that plugin and theme developers use WP_DEBUG
 * in their development environments.
 *
 * For information on other constants that can be used for debugging,
 * visit the Codex.
 *
 * @link https://codex.wordpress.org/Debugging_in_WordPress
 */
define('WP_DEBUG', false);

/* That's all, stop editing! Happy blogging. */

/** Absolute path to the WordPress directory. */
if ( !defined('ABSPATH') )
	define('ABSPATH', dirname(__FILE__) . '/');

/** Sets up WordPress vars and included files. */
require_once(ABSPATH . 'wp-settings.php');
