<?php

namespace Box\Spout3\Common\Helper;

/**
 * Class GlobalFunctionsHelper
 * This class wraps global functions to facilitate testing
 *
 * @codeCoverageIgnore
 */
class GlobalFunctionsHelper
{

    /**
     * @var array An array of current string buffers associated by resource ID.
     */
    static $buffers = [];

    /**
     * @var int The number of bytes to buffer fwrites to.
     */
    static $buffer_size = 1024 * 1024;

    /**
     * Wrapper around global function fopen()
     * @see fopen()
     *
     * @param string $fileName
     * @param string $mode
     * @return resource|bool
     */
    public function fopen($fileName, $mode)
    {
        return fopen($fileName, $mode);
    }

    /**
     * Wrapper around global function fopen(), with buffering enabled.
     * @see fopen()
     *
     * @param string $fileName
     * @param string $mode
     * @return resource|bool
     */
    public static function fopen_buffered($fileName, $mode)
    {
        $resource = fopen($fileName, $mode);

        stream_set_chunk_size($resource, self::$buffer_size); // Prevent out buffer from being broken up into multiple, smaller-sized chunks when we go to fwrite.

        return $resource;
    }

    /**
     * Wrapper around global function fgets()
     * @see fgets()
     *
     * @param resource $handle
     * @param int|null $length
     * @return string
     */
    public function fgets($handle, $length = null)
    {
        return fgets($handle, $length);
    }

    /**
     * Wrapper around global function fputs()
     * @see fputs()
     *
     * @param resource $handle
     * @param string $string
     * @return int
     */
    public function fputs($handle, $string)
    {
        return fputs($handle, $string);
    }

    /**
     * Wrapper around global function fflush()
     * @see fflush()
     *
     * @param resource $handle
     * @return bool
     */
    public function fflush($handle)
    {
        return fflush($handle);
    }

    /**
     * Wrapper around global function fseek()
     * @see fseek()
     *
     * @param resource $handle
     * @param int $offset
     * @return int
     */
    public function fseek($handle, $offset)
    {
        return fseek($handle, $offset);
    }

    /**
     * Wrapper around global function fgetcsv()
     * @see fgetcsv()
     *
     * @param resource $handle
     * @param int|null $length
     * @param string|null $delimiter
     * @param string|null $enclosure
     * @return array
     */
    public function fgetcsv($handle, $length = null, $delimiter = null, $enclosure = null)
    {
        // PHP uses '\' as the default escape character. This is not RFC-4180 compliant...
        // To fix that, simply disable the escape character.
        // @see https://bugs.php.net/bug.php?id=43225
        // @see http://tools.ietf.org/html/rfc4180
        $escapeCharacter = "\0";

        return fgetcsv($handle, $length, $delimiter, $enclosure, $escapeCharacter);
    }

    /**
     * Wrapper around global function fputcsv()
     * @see fputcsv()
     *
     * @param resource $handle
     * @param array $fields
     * @param string|null $delimiter
     * @param string|null $enclosure
     * @return int
     */
    public function fputcsv($handle, array $fields, $delimiter = null, $enclosure = null)
    {
        // PHP uses '\' as the default escape character. This is not RFC-4180 compliant...
        // To fix that, simply disable the escape character.
        // @see https://bugs.php.net/bug.php?id=43225
        // @see http://tools.ietf.org/html/rfc4180
        $escapeCharacter = "\0";

        return fputcsv($handle, $fields, $delimiter, $enclosure, $escapeCharacter);
    }

    /**
     * Wrapper around global function fwrite()
     * @see fwrite()
     *
     * @param resource $handle
     * @param string $string
     * @return int
     */
    public function fwrite($handle, $string)
    {

        return fwrite($handle, $string);

    }


    /**
     * Wrapper around global function fwrite() that enables buffering
     * @see fwrite()
     *
     * @param resource $handle
     * @param string $string
     * @return int
     */
    public static function fwrite_buffered($handle, $string)
    {

        if (strlen($string) > self::$buffer_size) {
            return fwrite($handle, (self::$buffers[intval($handle)] ?? '') . $string);
        }

        else if (!isset(self::$buffers[intval($handle)])) {
            self::$buffers[intval($handle)] = $string;
        }

        else {
            self::$buffers[intval($handle)] .= $string;

            if (strlen(self::$buffers[intval($handle)]) > self::$buffer_size) {
                $res = fwrite($handle, self::$buffers[intval($handle)]);

                unset(self::$buffers[intval($handle)]);

                return $res;
            }
        }

        return 0;

    }

    /**
     * Wrapper around global function fclose()
     * @see fclose()
     *
     * @param resource $handle
     * @return bool
     */
    public function fclose($handle)
    {
        return fclose($handle);
    }

    /**
     * Wrapper around global function fclose() that clears appropriate buffers from {@see self::fwrite_buffered()}
     * @see fclose()
     *
     * @param resource $handle
     * @return bool
     */
    public static function fclose_buffered($handle)
    {

        if (isset(self::$buffers[intval($handle)])) {
            fwrite($handle, self::$buffers[intval($handle)]);
        }

        return fclose($handle);

    }

    /**
     * Wrapper around global function rewind()
     * @see rewind()
     *
     * @param resource $handle
     * @return bool
     */
    public function rewind($handle)
    {
        return rewind($handle);
    }

    /**
     * Wrapper around global function file_exists()
     * @see file_exists()
     *
     * @param string $fileName
     * @return bool
     */
    public function file_exists($fileName)
    {
        return file_exists($fileName);
    }

    /**
     * Wrapper around global function file_get_contents()
     * @see file_get_contents()
     *
     * @param string $filePath
     * @return string
     */
    public function file_get_contents($filePath)
    {
        $realFilePath = $this->convertToUseRealPath($filePath);

        return file_get_contents($realFilePath);
    }

    /**
     * Updates the given file path to use a real path.
     * This is to avoid issues on some Windows setup.
     *
     * @param string $filePath File path
     * @return string The file path using a real path
     */
    protected function convertToUseRealPath($filePath)
    {
        $realFilePath = $filePath;

        if ($this->isZipStream($filePath)) {
            if (preg_match('/zip:\/\/(.*)#(.*)/', $filePath, $matches)) {
                $documentPath = $matches[1];
                $documentInsideZipPath = $matches[2];
                $realFilePath = 'zip://' . realpath($documentPath) . '#' . $documentInsideZipPath;
            }
        } else {
            $realFilePath = realpath($filePath);
        }

        return $realFilePath;
    }

    /**
     * Returns whether the given path is a zip stream.
     *
     * @param string $path Path pointing to a document
     * @return bool TRUE if path is a zip stream, FALSE otherwise
     */
    protected function isZipStream($path)
    {
        return (strpos($path, 'zip://') === 0);
    }

    /**
     * Wrapper around global function feof()
     * @see feof()
     *
     * @param resource $handle
     * @return bool
     */
    public function feof($handle)
    {
        return feof($handle);
    }

    /**
     * Wrapper around global function is_readable()
     * @see is_readable()
     *
     * @param string $fileName
     * @return bool
     */
    public function is_readable($fileName)
    {
        return is_readable($fileName);
    }

    /**
     * Wrapper around global function basename()
     * @see basename()
     *
     * @param string $path
     * @param string|null $suffix
     * @return string
     */
    public function basename($path, $suffix = null)
    {
        return basename($path, $suffix);
    }

    /**
     * Wrapper around global function header()
     * @see header()
     *
     * @param string $string
     * @return void
     */
    public function header($string)
    {
        header($string);
    }

    /**
     * Wrapper around global function ob_end_clean()
     * @see ob_end_clean()
     *
     * @return void
     */
    public function ob_end_clean()
    {
        if (ob_get_length() > 0) {
            ob_end_clean();
        }
    }

    /**
     * Wrapper around global function iconv()
     * @see iconv()
     *
     * @param string $string The string to be converted
     * @param string $sourceEncoding The encoding of the source string
     * @param string $targetEncoding The encoding the source string should be converted to
     * @return string|bool the converted string or FALSE on failure.
     */
    public function iconv($string, $sourceEncoding, $targetEncoding)
    {
        return iconv($sourceEncoding, $targetEncoding, $string);
    }

    /**
     * Wrapper around global function mb_convert_encoding()
     * @see mb_convert_encoding()
     *
     * @param string $string The string to be converted
     * @param string $sourceEncoding The encoding of the source string
     * @param string $targetEncoding The encoding the source string should be converted to
     * @return string|bool the converted string or FALSE on failure.
     */
    public function mb_convert_encoding($string, $sourceEncoding, $targetEncoding)
    {
        return mb_convert_encoding($string, $targetEncoding, $sourceEncoding);
    }

    /**
     * Wrapper around global function stream_get_wrappers()
     * @see stream_get_wrappers()
     *
     * @return array
     */
    public function stream_get_wrappers()
    {
        return stream_get_wrappers();
    }

    /**
     * Wrapper around global function function_exists()
     * @see function_exists()
     *
     * @param string $functionName
     * @return bool
     */
    public function function_exists($functionName)
    {
        return function_exists($functionName);
    }
}
