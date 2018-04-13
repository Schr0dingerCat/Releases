package util

import (
	"archive/tar"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
)

func IsExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	} else {
		if os.IsNotExist(err) {
			return false, nil
		}
	}
	return false, err
}

func ListDir(dirPath string, suffix string) (files []string, err error) {
	files = make([]string, 0, 10)
	pthSep := string(os.PathSeparator)
	dirPath = strings.TrimSuffix(dirPath, pthSep)
	dir, err := ioutil.ReadDir(dirPath)
	if err != nil {
		return nil, err
	}
	suffix = strings.ToUpper(suffix)
	for _, fi := range dir {
		// if fi.IsDir() {
		// 	continue
		// }
		if strings.HasSuffix(strings.ToUpper(fi.Name()), suffix) {
			files = append(files, dirPath+pthSep+fi.Name())
		}
	}
	return files, nil
}

func WalkDir1(dirPath string, suffix string) (files []string, err error) {
	files = make([]string, 0, 20)
	pthSep := string(os.PathSeparator)
	dirPath = strings.TrimSuffix(dirPath, pthSep)
	suffix = strings.ToUpper(suffix)
	err = filepath.Walk(dirPath, func(filename string, fi os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if fi.IsDir() {
			return nil
		}
		if strings.HasSuffix(strings.ToUpper(fi.Name()), suffix) {
			files = append(files, filename)
		}
		return nil
	})
	return files, err
}

func WalkDir(dirPath string, suffix string) (files []string, infos []os.FileInfo, err error) {
	files = make([]string, 0, 20)
	infos = make([]os.FileInfo, 0, 20)
	pthSep := string(os.PathSeparator)
	dirPath = strings.TrimSuffix(dirPath, pthSep)
	suffix = strings.ToUpper(suffix)
	err = filepath.Walk(dirPath, func(filename string, fi os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if fi.IsDir() {
			return nil
		}
		if strings.HasSuffix(strings.ToUpper(fi.Name()), suffix) {
			files = append(files, filename)
			infos = append(infos, fi)
		}
		return nil
	})
	return files, infos, err
}

////////////tar
/***生成***/
func CreateTar(filesource, filetarget string, deleteIfExist bool) error {
	//create tar file with targetfile name
	tarfile, err := os.Create(filetarget)
	if err != nil {
		// if file is exist then delete file
		if err == os.ErrExist {
			if err := os.Remove(filetarget); err != nil {
				fmt.Println(err)
				return err
			}
		} else {
			fmt.Println(err)
			return err
		}
	}
	defer tarfile.Close()
	tarwriter := tar.NewWriter(tarfile)
	sfileInfo, err := os.Stat(filesource)
	if err != nil {
		fmt.Println(err)
		return err
	}
	if !sfileInfo.IsDir() {
		return tarFile(filetarget, filesource, sfileInfo, tarwriter)
	} else {
		return tarFolder(filesource, tarwriter)
	}
	return nil
}
func tarFile(directory string, filesource string, sfileInfo os.FileInfo, tarwriter *tar.Writer) error {
	sfile, err := os.Open(filesource)
	if err != nil {
		panic(err)
		return err
	}
	defer sfile.Close()
	header, err := tar.FileInfoHeader(sfileInfo, "")
	if err != nil {
		fmt.Println(err)
		return err
	}
	header.Name = directory
	err = tarwriter.WriteHeader(header)
	if err != nil {
		fmt.Println(err)
		return err
	}
	// can use buffer to copy the file to tar writer
	//    buf := make([]byte,15)
	//    if _, err = io.CopyBuffer(tarwriter, sfile, buf); err != nil {
	//        panic(err)
	//        return err
	//    }
	if _, err = io.Copy(tarwriter, sfile); err != nil {
		fmt.Println(err)
		panic(err)
		return err
	}
	return nil
}
func tarFolder(directory string, tarwriter *tar.Writer) error {
	var baseFolder string = filepath.Base(directory)
	//fmt.Println(baseFolder)
	return filepath.Walk(directory, func(targetpath string, file os.FileInfo, err error) error {
		//read the file failure
		if file == nil {
			panic(err)
			return err
		}
		if file.IsDir() {
			// information of file or folder
			header, err := tar.FileInfoHeader(file, "")
			if err != nil {
				return err
			}
			header.Name = filepath.Join(baseFolder, strings.TrimPrefix(targetpath, directory))
			fmt.Println(header.Name)
			if err = tarwriter.WriteHeader(header); err != nil {
				return err
			}
			os.Mkdir(strings.TrimPrefix(baseFolder, file.Name()), os.ModeDir)
			return nil
		} else {
			//baseFolder is the tar file path
			var fileFolder = filepath.Join(baseFolder, strings.TrimPrefix(targetpath, directory))
			return tarFile(fileFolder, targetpath, file, tarwriter)
		}
	})
}
